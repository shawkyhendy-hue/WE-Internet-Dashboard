import json
import os
import re
import hashlib
from datetime import datetime
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError

# Live WE portal connector.
# The connector logs in with Service Number -> Internet -> Password, listens to
# authenticated JSON/XHR traffic, then maps common WE response fields into the
# dashboard schema. It also has a DOM-text fallback for quota/usage/bill values.

FIELD_ALIASES = {
    'landline': ['serviceNumber','service_number','serviceNo','service_no','fixedLineNumber','fixed_line_number','landline','telephoneNumber','phoneNumber','number'],
    'package': ['packageName','package_name','bundleName','bundle_name','planName','plan_name','tariffName','tariff_name','offerName','offer_name','offeringName','offering_name','productName','product_name','mainOfferName','main_offer_name','commercialName','package','bundle','plan','offer'],
    'speed_mbps': ['speedMbps','speed_mbps','internetSpeed','downloadSpeed','download_speed','bandwidth','bandwidthMbps','speed'],
    'quota_gb': ['totalQuota','total_quota','quotaGB','quota_gb','quota','totalVolume','total_volume','totalAllowance','total_allowance','allowance','capacity','initialBalance','totalBalance','totalUnits','bundleQuota'],
    'used_gb': ['usedQuota','used_quota','consumedQuota','consumed_quota','usedGB','used_gb','consumedGB','consumed_gb','usage','usedVolume','consumedVolume','consumed','used'],
    'remaining_gb': ['remainingQuota','remaining_quota','remainingGB','remaining_gb','remainingVolume','remaining_volume','remainingBalance','availableBalance','remainingUnits','remaining','availableQuota','available'],
    'usage_pct': ['usagePercentage','usage_percentage','usagePercent','usage_percent','usagePct','usage_pct','consumptionPercentage','percentage'],
    'bill_egp': ['billAmount','bill_amount','invoiceAmount','invoice_amount','currentBill','current_bill','totalBill','monthlyBill','invoiceTotal','billValue'],
    'renewal_cost_egp': ['renewalCost','renewal_cost','renewCost','planRenewalCost'],
    'outstanding_egp': ['outstandingAmount','outstanding_amount','dueAmount','due_amount','balanceDue','balance_due','amountDue','unpaidAmount','outstandingBalance','currentBalance'],
    'due_date': ['remainingDays','remaining_days','daysRemaining','days_remaining','dueText','due_text'],
    'renewal_date': ['renewalDate','renewal_date','nextRenewalDate','next_renewal_date','expiryDate','expirationDate','bundleExpiryDate','validityEndDate','endDate','expireDate'],
    'status': ['serviceStatus','service_status','lineStatus','line_status','status'],
}


def _norm_key(k):
    return re.sub(r'[^a-z0-9]', '', str(k).lower())

_ALIAS_MAP = {field: {_norm_key(x) for x in aliases} for field, aliases in FIELD_ALIASES.items()}


def _first_visible(page, selectors):
    for selector in selectors:
        try:
            loc = page.locator(selector)
            if loc.count() and loc.first.is_visible():
                return loc.first
        except Exception:
            pass
    return None


def _select_service_type(page, service_type='Internet'):
    """Select WE Service Type reliably, including the Ant Design control used by my.te.eg."""
    service_type = (service_type or 'Internet').strip() or 'Internet'

    # 1) Native <select> fallback.
    selects = page.locator('select')
    for i in range(selects.count()):
        sel = selects.nth(i)
        try:
            labels = sel.locator('option').all_text_contents()
            if any(service_type.lower() in (x or '').lower() for x in labels):
                try:
                    sel.select_option(label=service_type)
                except Exception:
                    values = sel.locator('option').evaluate_all('(opts) => opts.map(o => ({text:o.textContent.trim(), value:o.value}))')
                    match = next((x for x in values if service_type.lower() in x['text'].lower()), None)
                    if match:
                        sel.select_option(match['value'])
                return
        except Exception:
            pass

    # 2) WE currently uses Ant Design. The real clickable target is the
    #    .ant-select-selector ancestor, not the placeholder span/search input.
    ant_input = page.locator('input#login_input_type_01, input[aria-controls="login_input_type_01_list"]').first
    if ant_input.count():
        try:
            ant_select = ant_input.locator('xpath=ancestor::div[contains(@class,"ant-select")][1]')
            selector = ant_select.locator('.ant-select-selector').first
            if selector.count():
                selector.scroll_into_view_if_needed()
                selector.click(force=True, timeout=5000)
            else:
                ant_select.click(force=True, timeout=5000)
            page.wait_for_timeout(250)

            # Prefer Ant option nodes so a duplicate "Internet" elsewhere cannot be clicked.
            option = page.locator('.ant-select-dropdown:not(.ant-select-dropdown-hidden) .ant-select-item-option').filter(has_text=service_type).first
            if not option.count():
                option = page.locator('.ant-select-item-option').filter(has_text=service_type).first
            if option.count():
                option.scroll_into_view_if_needed()
                option.click(force=True, timeout=5000)
                page.wait_for_timeout(250)
                return

            # Keyboard fallback once the combobox is open.
            ant_input.press('i')
            page.wait_for_timeout(150)
            ant_input.press('Enter')
            page.wait_for_timeout(250)
            val = ant_select.inner_text(timeout=2000)
            if service_type.lower() in val.lower():
                return
        except Exception:
            pass

    # 3) Generic combobox fallback. Click the nearest actual select container.
    candidates = [
        page.locator('.ant-select').filter(has_text='Select Type').first,
        page.locator('[role="combobox"]').first,
        page.get_by_text('Select Type', exact=True).first,
    ]
    trigger = None
    for c in candidates:
        try:
            if c.count() and c.is_visible():
                trigger = c
                break
        except Exception:
            pass
    if not trigger:
        raise RuntimeError('WE Service Type selector was not detected.')

    try:
        if 'ant-select' in (trigger.get_attribute('class') or ''):
            clickable = trigger.locator('.ant-select-selector').first
            (clickable if clickable.count() else trigger).click(force=True, timeout=5000)
        else:
            trigger.click(force=True, timeout=5000)
    except Exception as e:
        raise RuntimeError(f'WE Service Type selector could not be opened: {e}')

    page.wait_for_timeout(250)
    option = page.locator('.ant-select-item-option').filter(has_text=service_type).first
    if option.count():
        option.click(force=True, timeout=5000)
        return
    role_option = page.get_by_role('option', name=service_type, exact=True)
    if role_option.count():
        role_option.first.click(force=True, timeout=5000)
        return
    raise RuntimeError(f'WE service type option "{service_type}" was not found.')

def _as_number(value):
    if value is None or isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    s = str(value).replace(',', '').strip()
    m = re.search(r'-?\d+(?:\.\d+)?', s)
    return float(m.group()) if m else None


def _gb(value):
    n = _as_number(value)
    if n is None:
        return None
    s = str(value).lower()
    if 'mb' in s and 'mbps' not in s:
        return n / 1024.0
    if 'kb' in s and 'kbps' not in s:
        return n / 1024.0 / 1024.0
    return n


def _account_service_type(account):
    value = str((account or {}).get('portal_service_type') or 'Internet').strip() or 'Internet'
    return 'LandLine' if value.lower() == 'landline' else 'Internet'


def _pick(obj, field):
    if not isinstance(obj, dict):
        return None
    aliases = _ALIAS_MAP[field]
    for k, v in obj.items():
        if _norm_key(k) in aliases and v not in (None, '', [], {}):
            return v
    return None


def _walk_dicts(node):
    if isinstance(node, dict):
        yield node
        for v in node.values():
            yield from _walk_dicts(v)
    elif isinstance(node, list):
        for v in node:
            yield from _walk_dicts(v)


def _looks_like_line(obj):
    land = _pick(obj, 'landline')
    pkg = _pick(obj, 'package')
    quota = _pick(obj, 'quota_gb')
    used = _pick(obj, 'used_gb')
    bill = _pick(obj, 'bill_egp') or _pick(obj, 'outstanding_egp')
    status = _pick(obj, 'status')
    return bool(land or (pkg and (quota is not None or used is not None)) or (bill is not None and status is not None))


def _canonical(obj, account):
    quota = _gb(_pick(obj, 'quota_gb'))
    used = _gb(_pick(obj, 'used_gb'))
    remaining = _gb(_pick(obj, 'remaining_gb'))
    usage_pct = _as_number(_pick(obj, 'usage_pct'))
    if usage_pct is None and quota and used is not None:
        usage_pct = used / quota * 100.0
    if remaining is None and quota is not None and used is not None:
        remaining = max(0.0, quota - used)
    landline = str(_pick(obj, 'landline') or account.get('login_username') or '').strip()
    status = str(_pick(obj, 'status') or 'Active').strip()
    status_low = status.lower()
    if usage_pct is not None and usage_pct >= 90 and status_low in ('active','normal','enabled',''):
        status = 'High Usage'
    return {
        'site': account.get('site') or '',
        'account_number': account.get('account_number') or landline,
        'landline': landline,
        'service_type': _account_service_type(account),
        'package': str(_pick(obj, 'package') or '').strip(),
        'speed_mbps': int(_as_number(_pick(obj, 'speed_mbps')) or 0),
        'quota_gb': round(quota or 0, 3),
        'used_gb': round(used or 0, 3),
        'usage_pct': round(usage_pct or 0, 2),
        'remaining_gb': round(remaining or 0, 3),
        'bill_egp': round(_as_number(_pick(obj, 'bill_egp')) or 0, 2),
        'renewal_cost_egp': round(_as_number(_pick(obj, 'renewal_cost_egp')) or 0, 2),
        'outstanding_egp': round(_as_number(_pick(obj, 'outstanding_egp')) or 0, 2),
        'due_date': str(_pick(obj, 'due_date') or '').strip(),
        'renewal_date': str(_pick(obj, 'renewal_date') or '').strip(),
        'status': status or 'Active',
    }


def _merge_lines(lines, account):
    merged = {}
    for line in lines:
        key = re.sub(r'\D', '', str(line.get('landline') or '')) or str(account.get('login_username') or '')
        if not key:
            continue
        cur = merged.setdefault(key, {
            'site': account.get('site') or '', 'account_number': account.get('account_number') or key,
            'landline': line.get('landline') or account.get('login_username') or '', 'service_type': _account_service_type(account),
            'package':'','speed_mbps':0,'quota_gb':0,'used_gb':0,'usage_pct':0,'remaining_gb':0,
            'bill_egp':0,'renewal_cost_egp':0,'outstanding_egp':0,'due_date':'','renewal_date':'','status':'Active'
        })
        for k, v in line.items():
            if k in ('site','account_number','landline','service_type'):
                if v: cur[k] = v
            elif v not in (None, '', 0, 0.0):
                cur[k] = v
    return list(merged.values())


def _extract_from_json(payloads, account):
    lines = []
    for payload in payloads:
        for obj in _walk_dicts(payload):
            if _looks_like_line(obj):
                lines.append(_canonical(obj, account))
    return _merge_lines(lines, account)


def _extract_from_dom(text, account):
    # Conservative fallback: create the service line and parse obvious labelled values.
    line = {
        'site': account.get('site') or '', 'account_number': account.get('account_number') or '',
        'landline': account.get('login_username') or '', 'service_type': _account_service_type(account), 'package':'',
        'speed_mbps':0,'quota_gb':0,'used_gb':0,'usage_pct':0,'remaining_gb':0,'bill_egp':0,'renewal_cost_egp':0,
        'outstanding_egp':0,'due_date':'','renewal_date':'','status':'Active'
    }
    patterns = {
        'usage_pct': [r'(?:usage|consumption)[^\d]{0,20}(\d+(?:\.\d+)?)\s*%'],
        'remaining_gb': [r'(?:remaining|available)[^\d]{0,30}(\d+(?:\.\d+)?)\s*GB'],
        'quota_gb': [r'(?:quota|allowance|package)[^\d]{0,35}(\d+(?:\.\d+)?)\s*GB'],
        'bill_egp': [r'(?:bill|invoice|current bill)[^\d]{0,40}(\d+(?:\.\d+)?)\s*(?:EGP|LE|L\.E)'],
        'renewal_cost_egp': [r'(?:renewal cost|subscription fee|monthly fee)[^\d]{0,40}(\d+(?:\.\d+)?)\s*(?:EGP|LE|L\.E)'],
        'outstanding_egp': [r'(?:outstanding|amount due|due amount|current balance|balance)[^\d]{0,40}(\d+(?:\.\d+)?)\s*(?:EGP|LE|L\.E)'],
    }
    low = text or ''
    for field, pats in patterns.items():
        for pat in pats:
            m = re.search(pat, low, flags=re.I|re.S)
            if m:
                line[field] = float(m.group(1)); break
    if line['quota_gb'] and line['remaining_gb'] and not line['used_gb']:
        line['used_gb'] = max(0, line['quota_gb'] - line['remaining_gb'])
    if line['quota_gb'] and line['used_gb'] and not line['usage_pct']:
        line['usage_pct'] = line['used_gb'] / line['quota_gb'] * 100
    return [line]




def _flatten_scalars(node, prefix='', out=None, depth=0):
    if out is None:
        out = []
    if depth > 10 or len(out) > 10000:
        return out
    if isinstance(node, dict):
        for k, v in node.items():
            path = f"{prefix}.{k}" if prefix else str(k)
            if isinstance(v, (dict, list)):
                _flatten_scalars(v, path, out, depth + 1)
            else:
                out.append((path, v))
    elif isinstance(node, list):
        for i, v in enumerate(node[:100]):
            _flatten_scalars(v, f"{prefix}[{i}]", out, depth + 1)
    return out


def _extract_fields_by_path(payloads):
    """Best-effort extraction when WE nests values under non-standard keys."""
    found = {}
    rules = {
        'package': ('package','bundle','plan','tariff','offer','offering','product'),
        'speed_mbps': ('speed','bandwidth','downloadrate','downstream'),
        'quota_gb': ('totalquota','quota','totalvolume','allowance','capacity','totalbalance','totalunits'),
        'used_gb': ('usedquota','consumedquota','usedgb','consumedgb','usedvolume','consumedvolume','consumed','used'),
        'remaining_gb': ('remainingquota','remaininggb','remainingvolume','remainingbalance','availablebalance','remainingunits','availablequota','remaining'),
        'usage_pct': ('usagepercentage','usagepercent','usagepct','consumptionpercentage','percent'),
        'bill_egp': ('billamount','invoiceamount','currentbill','totalbill','monthlybill','invoicetotal','billvalue'),
        'renewal_cost_egp': ('renewalcost','renew_cost','planrenewalcost'),
        'outstanding_egp': ('outstandingamount','dueamount','balancedue','amountdue','unpaidamount','outstandingbalance'),
        'due_date': ('remainingdays','daysremaining','duetext'),
        'renewal_date': ('renewaldate','nextrenewaldate','expirydate','expirationdate','validityenddate','expiredate'),
        'status': ('servicestatus','linestatus'),
    }
    for payload in payloads:
        for path, value in _flatten_scalars(payload):
            if value in (None, '', [], {}):
                continue
            np = _norm_key(path)
            for field, needles in rules.items():
                if field in found:
                    continue
                if any(n in np for n in needles):
                    if field == 'package':
                        # Avoid selecting generic labels/booleans/IDs as package names.
                        sv = str(value).strip()
                        if not sv or sv.lower() in ('true','false','0','1','none') or (sv.isdigit() and len(sv) > 6):
                            continue
                    found[field] = value
    return found


def _merge_partial_line(base, extra):
    out = dict(base or {})
    for k, v in (extra or {}).items():
        if v not in (None, '', 0, 0.0):
            out[k] = v
    return out


def _extract_dom_rich(text, account):
    """Parse WE's authenticated Overview / Usage pages using the labels currently shown by my.te.eg.

    DOM values are intentionally preferred for the fields below because the portal can emit
    several unrelated quota/balance values in JSON. The visible labels are the user-facing
    source of truth for the selected Internet service.
    """
    line = _extract_from_dom(text, account)[0]
    txt = text or ''
    is_landline = _account_service_type(account).lower() == 'landline'
    line['service_type'] = 'LandLine' if is_landline else 'Internet'

    # Current Plan\nSuper speed 2- (750GB)
    m = re.search(r'Current\s+Plan\s*\n\s*([^\n]{2,120})', txt, flags=re.I)
    if m:
        line['package'] = m.group(1).strip(' :-')[:120]
    else:
        for pat in [
            r'(?:package|current package|plan|bundle|main plan)\s*[:\-]?\s*\n?\s*([^\n]{2,100})',
            r'(?:package name|plan name|bundle name)\s*[:\-]?\s*([^\n]{2,100})',
        ]:
            m = re.search(pat, txt, flags=re.I)
            if m:
                val = m.group(1).strip(' :-')
                if val and not re.search(r'^(usage|remaining|quota|bill|balance)$', val, re.I):
                    line['package'] = val[:120]
                    break

    # WE Usage page currently renders:
    #   Main Quota
    #   579.43 GB / 750 GB
    # The first value is the remaining amount shown on the main quota bar and the
    # second value is the total quota. Prefer this pair over generic JSON balances.
    mq = re.search(
        r'Main\s+Quota[\s\S]{0,220}?(\d+(?:\.\d+)?)\s*GB\s*/\s*(\d+(?:\.\d+)?)\s*GB',
        txt, flags=re.I
    )
    if mq:
        line['remaining_gb'] = float(mq.group(1))
        line['quota_gb'] = float(mq.group(2))

    # The circular usage widget renders the consumed amount as "180.57 Used".
    # Prefer that explicit widget value. Fall back to "Used 170.57 GB" if needed.
    used = re.search(r'(\d+(?:\.\d+)?)\s*Used\b', txt, flags=re.I)
    if used:
        line['used_gb'] = float(used.group(1))
    else:
        used = re.search(r'\bUsed\s*(\d+(?:\.\d+)?)\s*GB', txt, flags=re.I)
        if used:
            line['used_gb'] = float(used.group(1))

    # If Main Quota was not found, use more generic quota/remaining labels.
    if not line.get('quota_gb'):
        for pat in [
            r'(?:total quota|quota|allowance|total allowance|capacity)\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*GB',
            r'(\d+(?:\.\d+)?)\s*GB\s*(?:quota|allowance|total)',
        ]:
            m = re.search(pat, txt, flags=re.I)
            if m:
                line['quota_gb'] = float(m.group(1)); break
    if not line.get('remaining_gb'):
        for pat in [
            r'(?:remaining|available)\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*GB',
            r'(\d+(?:\.\d+)?)\s*GB?\s*Remaining\b',
        ]:
            m = re.search(pat, txt, flags=re.I)
            if m:
                line['remaining_gb'] = float(m.group(1)); break

    # Usage percentage is derived from Used / Total Quota because that is what the
    # dashboard needs, even when WE does not display a percentage in the page.
    if line.get('quota_gb') and line.get('used_gb') is not None:
        line['usage_pct'] = line['used_gb'] / line['quota_gb'] * 100.0

    # Overview page renders "Renewal Cost: 925 EGP". This is NOT a bill.
    rc = re.search(r'Renewal\s+Cost\s*:\s*(\d+(?:\.\d+)?)\s*EGP', txt, flags=re.I)
    if rc:
        line['renewal_cost_egp'] = float(rc.group(1))

    # Overview page renders "Renewal Date: 21-09-2026, 9 Remaining Days".
    rd = re.search(r'Renewal\s+Date\s*:\s*(\d{1,2}[\-/]\d{1,2}[\-/]\d{4})\s*,?\s*(\d+)\s+Remaining\s+Days', txt, flags=re.I)
    if rd:
        line['renewal_date'] = rd.group(1)
        line['due_date'] = f"{rd.group(2)} Remaining Days"
    else:
        rd = re.search(r'Renewal\s+Date\s*:\s*(\d{1,2}[\-/]\d{1,2}[\-/]\d{4})', txt, flags=re.I)
        if rd:
            line['renewal_date'] = rd.group(1)
        dd = re.search(r'(\d+)\s+Remaining\s+Days', txt, flags=re.I)
        if dd:
            line['due_date'] = f"{dd.group(1)} Remaining Days"

    # Home page shows "Current Balance 2570.01 EGP"; treat this as Outstanding.
    cb = re.search(r'Current\s+Balance[\s\S]{0,80}?(\d+(?:\.\d+)?)\s*EGP', txt, flags=re.I)
    if cb:
        line['outstanding_egp'] = float(cb.group(1))

    # Optional speed value, only when the page explicitly states Mbps.
    m = re.search(r'(?:internet speed|download speed|speed)\s*[:\-]?\s*(\d+(?:\.\d+)?)\s*(?:Mbps|Mb/s)', txt, flags=re.I)
    if m:
        line['speed_mbps'] = float(m.group(1))

    # WE LandLine accounts do not expose Internet quota widgets. Capture the
    # visible plan/bill labels from the authenticated LandLine pages instead.
    if is_landline:
        if not line.get('package'):
            for pat in [
                r'(?:current plan|plan|package|tariff|service)\s*[:\-]?\s*\n?\s*([^\n]{2,100})',
                r'(?:landline plan|fixed line plan)\s*[:\-]?\s*([^\n]{2,100})',
            ]:
                mm=re.search(pat,txt,flags=re.I)
                if mm:
                    val=mm.group(1).strip(' :-')
                    if val and not re.search(r'^(bill|balance|status|landline)$',val,re.I):
                        line['package']=val[:120]; break
        if not line.get('renewal_cost_egp'):
            mm=re.search(r'(?:renewal cost|monthly fee|subscription fee|package fee)[^\d]{0,45}([0-9][0-9,]*(?:\.[0-9]+)?)\s*(?:EGP|LE|L\.E)',txt,flags=re.I|re.S)
            if mm: line['renewal_cost_egp']=float(mm.group(1).replace(',',''))
        if not line.get('bill_egp'):
            mm=re.search(r'(?:current bill|bill amount|invoice amount|bill)[^\d]{0,45}([0-9][0-9,]*(?:\.[0-9]+)?)\s*(?:EGP|LE|L\.E)',txt,flags=re.I|re.S)
            if mm: line['bill_egp']=float(mm.group(1).replace(',',''))
        # LandLine portal labels this value as "Paid in Advance". For the
        # dashboard, this is the requested Outstanding value/source of truth.
        mm=re.search(r'Paid\s+in\s+Advance[^\d]{0,60}([0-9][0-9,]*(?:\.[0-9]+)?)\s*(?:EGP|LE|L\.E|ج\.?\s?م)?',txt,flags=re.I|re.S)
        if not mm:
            mm=re.search(r'([0-9][0-9,]*(?:\.[0-9]+)?)\s*(?:EGP|LE|L\.E|ج\.?\s?م)?[^\n]{0,30}Paid\s+in\s+Advance',txt,flags=re.I|re.S)
        if mm:
            line['outstanding_egp']=float(mm.group(1).replace(',',''))
        elif not line.get('outstanding_egp'):
            mm=re.search(r'(?:outstanding|amount due|due amount|current balance|balance)[^\d]{0,45}([0-9][0-9,]*(?:\.[0-9]+)?)\s*(?:EGP|LE|L\.E)',txt,flags=re.I|re.S)
            if mm: line['outstanding_egp']=float(mm.group(1).replace(',',''))

    # Generic outstanding fallback. Do not confuse Renewal Cost with outstanding.
    if not line.get('outstanding_egp'):
        m = re.search(r'(?:outstanding|amount due|due amount|unpaid|outstanding balance)[^\d]{0,40}(\d+(?:\.\d+)?)\s*(?:EGP|LE|L\.E|ج\.م)?', txt, flags=re.I|re.S)
        if m:
            line['outstanding_egp'] = float(m.group(1))

    line['status'] = 'Active'
    return line


def _extract_overview_exact(text):
    """Extract the three user-facing renewal fields from WE #/overview only.

    These values override generic API/path mappings because the portal exposes
    unrelated numeric timestamps and balance fields elsewhere.
    """
    txt = text or ''
    out = {}
    m = re.search(r'Renewal\s+Cost\s*:\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*EGP', txt, flags=re.I)
    if m:
        out['renewal_cost_egp'] = float(m.group(1).replace(',', ''))
    m = re.search(r'Renewal\s+Date\s*:\s*(\d{1,2}[\-/]\d{1,2}[\-/]\d{4})(?:\s*,?\s*(\d+)\s+Remaining\s+Days)?', txt, flags=re.I)
    if m:
        out['renewal_date'] = m.group(1)
        if m.group(2):
            out['due_date'] = f"{m.group(2)} Remaining Days"
    if 'due_date' not in out:
        m = re.search(r'\b(\d+)\s+Remaining\s+Days\b', txt, flags=re.I)
        if m:
            out['due_date'] = f"{m.group(1)} Remaining Days"
    return out


def _goto_hash_and_text(page, route, timeout_ms=15000):
    """Navigate directly to an authenticated WE SPA hash route and return body text."""
    try:
        current = page.url or 'https://my.te.eg/echannel/'
        base = current.split('#', 1)[0]
        if not base.endswith('/'):
            base += '/'
        target = base + '#/' + route.lstrip('/')
        page.goto(target, wait_until='domcontentloaded', timeout=timeout_ms)
        try:
            page.wait_for_load_state('networkidle', timeout=10000)
        except Exception:
            pass
        page.wait_for_timeout(2200)
        return page.locator('body').inner_text(timeout=10000)
    except Exception:
        return ''


def _flatten_paths(node, prefix='', out=None, depth=0):
    if out is None:
        out = []
    if depth > 8 or len(out) >= 2000:
        return out
    if isinstance(node, dict):
        for k, v in node.items():
            key = str(k)
            path = f"{prefix}.{key}" if prefix else key
            if isinstance(v, (dict, list)):
                _flatten_paths(v, path, out, depth + 1)
            else:
                # Never copy obvious secrets/tokens into diagnostics.
                low = key.lower()
                if any(x in low for x in ['password','token','authorization','cookie','secret','session','jwt','access','refresh']):
                    val = '<redacted>'
                else:
                    text = str(v)
                    val = text[:180]
                out.append({'path': path, 'value': val, 'type': type(v).__name__})
    elif isinstance(node, list):
        for i, v in enumerate(node[:5]):
            _flatten_paths(v, f"{prefix}[{i}]", out, depth + 1)
    return out


def _save_diagnostics(account, responses, payloads, body_text, force=False, storage=None, resources=None):
    if not force and os.getenv('WE_DIAGNOSTIC_CAPTURE', '0').strip().lower() not in ('1','true','yes'):
        return None
    folder = Path(os.path.expandvars(os.path.expanduser(os.getenv('WE_DATA_DIR') or str(Path.home() / 'WE-Dashboard-Data')))) / 'diagnostics'
    folder.mkdir(exist_ok=True)
    safe = re.sub(r'\D', '', str(account.get('login_username') or 'account'))[-4:] or 'account'
    stamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    path = folder / f'we_mapping_{safe}_{stamp}.json'

    payload_info = []
    for idx, payload in enumerate(payloads[-120:]):
        flat = _flatten_paths(payload)
        interesting = []
        words = ('quota','usage','consum','remain','bundle','package','plan','speed','bill','invoice','due','renew','expiry','expire','balance','service','number','status','amount')
        for item in flat:
            if any(w in item['path'].lower() for w in words):
                interesting.append(item)
        payload_info.append({
            'index': idx,
            'root_type': type(payload).__name__,
            'interesting_fields': interesting[:200],
            'all_paths_sample': flat[:120],
        })

    # Keep only a short page text preview and mask long digit sequences.
    preview = (body_text or '')[:12000]
    preview = re.sub(r'(?<!\d)\d{7,}(?!\d)', '<number-redacted>', preview)
    data = {
        'timestamp': stamp,
        'service_suffix': safe,
        'responses': responses[-150:],
        'payloads': payload_info,
        'page_text_preview': preview,
        'storage': storage or {},
        'resource_urls': (resources or [])[-250:],
    }
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    return str(path)


def _has_meaningful_data(lines):
    for line in lines or []:
        if (line.get('package') or line.get('speed_mbps') or line.get('quota_gb') or
            line.get('used_gb') or line.get('remaining_gb') or line.get('usage_pct') or
            line.get('bill_egp') or line.get('outstanding_egp') or line.get('due_date') or
            line.get('renewal_date')):
            return True
    return False

class WEVerificationRequired(RuntimeError):
    pass


def _session_state_path(account):
    base = Path(os.path.expandvars(os.path.expanduser(os.getenv('WE_DATA_DIR') or str(Path.home() / 'WE-Dashboard-Data')))) / 'sessions'
    base.mkdir(parents=True, exist_ok=True)
    identity = str(account.get('account_number') or account.get('login_username') or 'we')
    safe = hashlib.sha256(identity.encode('utf-8')).hexdigest()[:16]
    return base / f'{safe}.json'


def fetch_we_lines(account, interactive=False):
    """Sync one WE account.

    Normal operation always starts headless and reuses the saved authenticated browser
    session. If WE requires CAPTCHA/verification, a manual Sync WE Now call retries once
    with visible Chrome. Scheduled hourly syncs never pop up a browser window.
    """
    try:
        return _fetch_we_lines_once(account, headless_override=True)
    except WEVerificationRequired:
        if not interactive:
            raise RuntimeError('WE verification is required. Click Sync WE Now to open Chrome and complete the verification code.')
        return _fetch_we_lines_once(account, headless_override=False)


def _fetch_we_lines_once(account, headless_override=None):
    service_number = (account.get('login_username') or '').strip()
    password = (account.get('password') or '').strip()
    service_type = (account.get('portal_service_type') or 'Internet').strip() or 'Internet'
    url = os.getenv('WE_LOGIN_URL', 'https://my.te.eg/')
    headless = (bool(headless_override) if headless_override is not None else os.getenv('WE_BROWSER_HEADLESS', '1').strip().lower() not in ('0','false','no'))
    wait_after_login = max(3, int(os.getenv('WE_WAIT_AFTER_LOGIN_SECONDS', '8')))

    if not service_number or not password:
        raise RuntimeError('WE Service Number or password is missing.')

    payloads, response_urls = [], []
    with sync_playwright() as p:
        launch_kwargs = {'headless': headless}
        # Prefer the locally installed Google Chrome. The WE portal may render a blank
        # /echannel page in Playwright's bundled Chromium on some Windows machines.
        if os.getenv('WE_USE_INSTALLED_CHROME', '1').strip().lower() not in ('0','false','no'):
            launch_kwargs['channel'] = 'chrome'
        try:
            browser = p.chromium.launch(**launch_kwargs)
        except Exception:
            launch_kwargs.pop('channel', None)
            browser = p.chromium.launch(**launch_kwargs)
        state_path = _session_state_path(account)
        context_kwargs = {'viewport': {'width':1440,'height':1000}, 'locale':'en-US'}
        if state_path.exists():
            try:
                context_kwargs['storage_state'] = str(state_path)
            except Exception:
                pass
        context = browser.new_context(**context_kwargs)
        page = context.new_page()

        def on_response(resp):
            try:
                from urllib.parse import urlsplit, urlunsplit
                u = urlsplit(resp.url)
                safe_url = urlunsplit((u.scheme, u.netloc, u.path, '', ''))
                ctype = (resp.headers.get('content-type') or '').lower()
                rtype = ''
                try:
                    rtype = resp.request.resource_type
                except Exception:
                    pass
                if rtype not in ('xhr','fetch') and 'json' not in ctype:
                    return
                response_urls.append({'url': safe_url, 'status': resp.status, 'content_type': ctype, 'resource_type': rtype})
                data = None
                try:
                    data = resp.json()
                except Exception:
                    try:
                        txt = resp.text()
                        stripped = (txt or '').lstrip()
                        if stripped.startswith('{') or stripped.startswith('['):
                            data = json.loads(stripped)
                    except Exception:
                        data = None
                if isinstance(data, (dict, list)):
                    payloads.append(data)
            except Exception:
                pass

        page.on('response', on_response)
        try:
            session_reused = False
            if state_path.exists():
                try:
                    # Test the saved authenticated session directly against the read-only Overview route.
                    base = url.rstrip('/')
                    if '/echannel' not in base:
                        base = base + '/echannel'
                    page.goto(base + '/#/overview', wait_until='domcontentloaded', timeout=60000)
                    page.wait_for_timeout(2500)
                    saved_body = page.locator('body').inner_text(timeout=10000).lower()
                    password_visible = page.locator('input[type="password"]:visible').count() > 0
                    session_reused = (not password_visible) and any(x in saved_body for x in [
                        'current plan', 'current balance', 'account overview', 'usage overview', 'manage another number'
                    ])
                    if session_reused:
                        print(f'WE saved session reused for {service_number}; background login skipped.')
                except Exception:
                    session_reused = False

            if not session_reused:
                page.goto(url, wait_until='domcontentloaded', timeout=60000)
                # WE redirects root to /echannel and its SPA can take several seconds to render.
                # Wait for the password field instead of inspecting the DOM immediately.
                try:
                    page.locator('input[type="password"]').first.wait_for(state='visible', timeout=30000)
                except PlaywrightTimeoutError:
                    # One reload recovers an occasional blank /echannel shell.
                    page.reload(wait_until='domcontentloaded', timeout=60000)
                    try:
                        page.locator('input[type="password"]').first.wait_for(state='visible', timeout=30000)
                    except PlaywrightTimeoutError:
                        title = page.title()
                        raise RuntimeError(f'WE login page stayed blank/not rendered. URL={page.url} Title={title!r}. Try normal Chrome and check proxy/security filtering.')

                service_el = _first_visible(page, [
                    'input[name*="service" i]', 'input[placeholder*="service" i]',
                    'input[type="tel"]', 'input[type="text"]'
                ])
                password_el = _first_visible(page, ['input[type="password"]'])
                if not service_el or not password_el:
                    raise RuntimeError(f'WE login fields were not detected at {page.url}. Portal layout may have changed.')

                service_el.fill(service_number)
                _select_service_type(page, service_type)
                password_el.fill(password)
                login = page.get_by_role('button', name='Login', exact=False)
                if not login.count():
                    login = page.locator('button[type="submit"]').first
                if not login.count():
                    raise RuntimeError('WE Login button was not detected.')
                login.click()

                # WE can require an interactive CAPTCHA/verification step after credentials.
                # Do not try to bypass it. Keep the visible Chrome window open and wait for
                # the user to complete the challenge manually, then continue automatically.
                captcha_timeout = max(60, int(os.getenv('WE_CAPTCHA_TIMEOUT_SECONDS', '240')))
                page.wait_for_timeout(1200)
                body = page.locator('body').inner_text(timeout=10000)
                body_low = body.lower()
                verification_markers = [
                    'verification', 'please enter the code below', 'captcha',
                    'security code', 'enter the code', 'رمز التحقق'
                ]
                captcha_seen = any(x in body_low for x in verification_markers)
                if captcha_seen:
                    if headless:
                        raise WEVerificationRequired('WE requires interactive verification/CAPTCHA.')
                    print(f'WE verification detected for {service_number}. Complete the CAPTCHA/code in the opened Chrome window. Waiting up to {captcha_timeout} seconds...')
                    deadline = __import__('time').time() + captcha_timeout
                    verified = False
                    while __import__('time').time() < deadline:
                        page.wait_for_timeout(1500)
                        try:
                            body = page.locator('body').inner_text(timeout=5000)
                        except Exception:
                            body = ''
                        low = body.lower()
                        still_verifying = any(x in low for x in verification_markers)
                        authenticated_hint = any(x in low for x in [
                            'account overview', 'usage', 'bills & payments',
                            'my account', 'current balance', 'remaining'
                        ])
                        password_visible = False
                        try:
                            password_visible = page.locator('input[type="password"]:visible').count() > 0
                        except Exception:
                            pass
                        if not still_verifying and (authenticated_hint or not password_visible):
                            verified = True
                            print('WE verification completed. Continuing automatic data capture...')
                            break
                    if not verified:
                        raise RuntimeError(f'WE verification was not completed within {captcha_timeout} seconds. Press Sync WE Now again and complete the code in the opened Chrome window.')

                try:
                    page.wait_for_load_state('networkidle', timeout=25000)
                except PlaywrightTimeoutError:
                    pass

                # Wait for authenticated API calls/widgets to populate.
                page.wait_for_timeout(wait_after_login * 1000)
                body = page.locator('body').inner_text(timeout=10000)
                body_low = body.lower()
                if ('sign in' in body_low or 'login' in body_low) and page.locator('input[type="password"]:visible').count():
                    if any(x in body_low for x in ['invalid','incorrect','wrong password','try again','not valid']):
                        raise RuntimeError('WE login was rejected. Check Service Number, Internet type and password.')

                    # Some WE sessions complete authentication after the verification dialog
                    # but leave the SPA on the sign-in route. Probe the authenticated Overview
                    # route in the same browser context before declaring failure.
                    try:
                        base = url.rstrip('/')
                        if '/echannel' not in base:
                            base = base + '/echannel'
                        page.goto(base + '/#/overview', wait_until='domcontentloaded', timeout=60000)
                        page.wait_for_timeout(4000)
                        probe = page.locator('body').inner_text(timeout=10000).lower()
                        password_visible = page.locator('input[type="password"]:visible').count() > 0
                        auth_ok = (not password_visible) and any(x in probe for x in [
                            'current plan','current balance','account overview','usage overview',
                            'manage another number','renewal cost','service number'
                        ])
                        if auth_ok:
                            print('WE authentication confirmed through Overview after verification.')
                        else:
                            raise RuntimeError('WE is still on the sign-in page after verification; authentication did not complete.')
                    except RuntimeError:
                        raise
                    except Exception:
                        raise RuntimeError('WE is still on the sign-in page after verification; authentication did not complete.')

            # Capture the authenticated pages explicitly. WE is a hash-routed SPA and
            # the visible menu label is not always "Account Overview". Direct hash
            # navigation is much more reliable than clicking text labels.
            page_texts = []
            home_text = ''
            try:
                home_text = page.locator('body').inner_text(timeout=10000)
                page_texts.append(home_text)
            except Exception:
                pass

            overview_text = _goto_hash_and_text(page, 'overview')
            if overview_text:
                page_texts.append(overview_text)

            usage_text = _goto_hash_and_text(page, 'usage')
            if usage_text:
                page_texts.append(usage_text)

            # Optional billing page. Keep as a passive read-only route; if WE changes
            # the route this simply returns an empty string and does not break sync.
            bills_text = _goto_hash_and_text(page, 'billsummary')
            if bills_text:
                page_texts.append(bills_text)

            body = '\n\n--- WE PAGE ---\n\n'.join(page_texts)
            lines = _extract_from_json(payloads, account)
            # Merge nested/path-based values that may live in different JSON objects.
            path_fields = _extract_fields_by_path(payloads)
            if lines:
                merged_line = _merge_partial_line(lines[0], _canonical(path_fields, account))
            else:
                merged_line = _canonical(path_fields, account)
            # Merge labelled DOM values from all three pages.
            dom_line = _extract_dom_rich(body, account)
            merged_line = _merge_partial_line(merged_line, dom_line)

            # Force exact user-facing Overview values over generic JSON fields.
            # This prevents timestamps such as 4133894400000 from replacing the
            # displayed Renewal Date and ensures Due is the displayed remaining days.
            overview_exact = _extract_overview_exact(overview_text)
            for key in ('renewal_cost_egp', 'due_date', 'renewal_date'):
                if key in overview_exact:
                    merged_line[key] = overview_exact[key]

            lines = [merged_line] if merged_line.get('landline') else []
            storage = {}
            try:
                storage['localStorage_keys'] = page.evaluate('Object.keys(localStorage)')
                storage['sessionStorage_keys'] = page.evaluate('Object.keys(sessionStorage)')
            except Exception:
                pass
            resources = []
            try:
                resources = page.evaluate('performance.getEntriesByType(\"resource\").map(r => r.name)')
                from urllib.parse import urlsplit, urlunsplit
                cleaned=[]
                for ru in resources:
                    try:
                        uu=urlsplit(str(ru)); cleaned.append(urlunsplit((uu.scheme,uu.netloc,uu.path,'','')))
                    except Exception:
                        pass
                resources=cleaned
            except Exception:
                resources=[]
            mapping_failed = not _has_meaningful_data(lines)

            def _present(v):
                return v not in (None, '', 0, 0.0)
            if service_type.lower() == 'landline':
                # LandLine has no Internet quota/usage requirement. A valid sync is based
                # on plan/billing/renewal information from the authenticated WE account.
                core_keys = ('package','bill_egp','renewal_cost_egp','outstanding_egp','due_date','renewal_date')
                core_count = sum(1 for k in core_keys if lines and _present(lines[0].get(k)))
                mapping_failed = not bool(lines)
                complete_mapping = bool(lines) and core_count >= 1
                partial_mapping = bool(lines) and core_count == 0
            else:
                # A WE Internet sync is considered complete when the core values visible
                # in Overview/Usage were captured. Speed and outstanding balance are optional.
                core_keys = ('package', 'quota_gb', 'used_gb', 'remaining_gb', 'renewal_cost_egp', 'due_date', 'renewal_date')
                core_count = sum(1 for k in core_keys if lines and _present(lines[0].get(k)))
                complete_mapping = bool(lines) and core_count >= 6 and _present(lines[0].get('quota_gb')) and _present(lines[0].get('remaining_gb'))
                partial_mapping = (not mapping_failed and not complete_mapping)

            diag_path = _save_diagnostics(account, response_urls, payloads, body, force=(mapping_failed or partial_mapping), storage=storage, resources=resources)
            if mapping_failed:
                print(f'WE mapping diagnostic saved to: {diag_path}')
                hint = f' Diagnostic file: {diag_path}' if diag_path else ''
                raise RuntimeError('Login completed, but WE service fields were not mapped from the current portal responses.' + hint)
            if partial_mapping:
                print(f'WE partial mapping ({core_count}/{len(core_keys)} core fields). Diagnostic saved to: {diag_path}')
                lines[0]['status'] = 'Partial Sync'
            else:
                # Successful mapping must explicitly replace the placeholder
                # Pending Sync state stored when the account was first added.
                lines[0]['status'] = 'High Usage' if float(lines[0].get('usage_pct') or 0) >= 90 else 'Active'
                print(f'WE mapping complete ({core_count}/{len(core_keys)} core fields). Status set to {lines[0]["status"]}.')
            try:
                context.storage_state(path=str(state_path))
            except Exception as e:
                print('WE session state could not be saved:', e)
            return lines
        finally:
            context.close()
            browser.close()
