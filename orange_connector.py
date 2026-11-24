import json
import os
import re
import hashlib
from pathlib import Path
from datetime import datetime, date
from playwright.sync_api import sync_playwright

ORANGE_LOGIN_URL = os.getenv('ORANGE_LOGIN_URL', 'https://www.orange.eg/ar/myaccount/login?ReturnUrl=%2Far%2F')
ORANGE_ACCOUNT_URL = os.getenv('ORANGE_ACCOUNT_URL', 'https://www.orange.eg/ar/myaccount/')
ORANGE_TARIFF_URL = os.getenv('ORANGE_TARIFF_URL', 'https://www.orange.eg/ar/myaccount/manage-tariff')
ORANGE_INTERNET_URL = os.getenv('ORANGE_INTERNET_URL', 'https://www.orange.eg/ar/myaccount/internet')
ORANGE_CAPTCHA_TIMEOUT_SECONDS = int(os.getenv('ORANGE_CAPTCHA_TIMEOUT_SECONDS', '240'))

ALIASES = {
    'package': ['packageName','bundleName','planName','tariffName','offerName','productName','package','bundle','plan','tariff','ratePlanName'],
    'quota_gb': ['totalQuota','quotaGB','quota','totalVolume','totalAllowance','allowance','capacity','bundleQuota','totalUnits','totalData','packageQuota'],
    'used_gb': ['usedQuota','consumedQuota','usedGB','consumedGB','usage','usedVolume','consumedVolume','consumed','used','usedData'],
    'remaining_gb': ['remainingQuota','remainingGB','remainingVolume','remainingBalance','availableBalance','remaining','availableQuota','available','remainingData','remainingUnits'],
    'usage_pct': ['usagePercentage','usagePercent','usagePct','consumptionPercentage','percentage','consumedPercentage'],
    'renewal_cost_egp': ['renewalCost','renewCost','planRenewalCost','monthlyFee','monthlyCharge','subscriptionFee','packagePrice','tariffPrice','renewalAmount','price'],
    'outstanding_egp': ['outstandingAmount','dueAmount','balanceDue','amountDue','unpaidAmount','outstandingBalance','currentBalance','balance','mainBalance','creditBalance'],
    'renewal_date': ['renewalDate','nextRenewalDate','expiryDate','expirationDate','bundleExpiryDate','validityEndDate','endDate','expireDate','dueDate','nextRenewal'],
    'due_date': ['remainingDays','daysRemaining','daysToRenewal','renewalIn','daysLeft'],
    'status': ['serviceStatus','lineStatus','status','subscriptionStatus'],
}

def _norm(k): return re.sub(r'[^a-z0-9]', '', str(k).lower())
AMAP = {k: {_norm(x) for x in v} for k, v in ALIASES.items()}

def _num(v):
    if v is None or isinstance(v, bool): return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).replace(',', '')
    m = re.search(r'-?\d+(?:\.\d+)?', s)
    return float(m.group()) if m else None

def _gb(v, key_hint=''):
    n = _num(v)
    if n is None: return None
    s = (str(v) + ' ' + key_hint).lower()
    if any(x in s for x in ('mb','mega','ميجا')) and 'mbps' not in s:
        return n / 1024.0
    if any(x in s for x in ('kb','كيلو')):
        return n / (1024.0 * 1024.0)
    return n

def _walk(node):
    if isinstance(node, dict):
        yield node
        for v in node.values(): yield from _walk(v)
    elif isinstance(node, list):
        for v in node: yield from _walk(v)

def _pick(obj, field):
    if not isinstance(obj, dict): return None, ''
    aliases = AMAP[field]
    for k, v in obj.items():
        if _norm(k) in aliases and v not in (None, '', [], {}):
            return v, str(k)
    return None, ''

def _data_root():
    return Path(os.path.expandvars(os.path.expanduser(os.getenv('WE_DATA_DIR') or str(Path.home() / 'WE-Dashboard-Data'))))

def _session_path(account):
    folder = _data_root() / 'sessions'; folder.mkdir(parents=True, exist_ok=True)
    identity = 'orange-mobile:' + str(account.get('login_username') or account.get('account_number') or 'orange')
    return folder / f"orange_mobile_{hashlib.sha256(identity.encode()).hexdigest()[:16]}.json"

def _scrub(x):
    if isinstance(x, dict):
        out = {}
        for k, v in x.items():
            kl = str(k).lower()
            if any(w in kl for w in ('password','token','cookie','authorization','session','secret','captcha')):
                out[k] = '[REDACTED]'
            else:
                out[k] = _scrub(v)
        return out
    if isinstance(x, list): return [_scrub(v) for v in x[:120]]
    if isinstance(x, str):
        # Mask long mobile-like values in diagnostics.
        return re.sub(r'(?<!\d)(01\d{9})(?!\d)', lambda m: '01******' + m.group(1)[-3:], x)
    return x

def _diag(account, urls, payloads, texts):
    folder = _data_root() / 'diagnostics'; folder.mkdir(parents=True, exist_ok=True)
    identity = re.sub(r'\D','',str(account.get('login_username') or account.get('account_number') or 'orange'))[-4:] or 'line'
    path = folder / f"orange_mobile_mapping_{identity}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    obj = {
        'provider': 'Orange Home Wireless',
        'generated_at': datetime.now().isoformat(timespec='seconds'),
        'response_urls': _scrub(urls[-140:]),
        'payloads': _scrub(payloads[-80:]),
        'page_text_preview': _scrub('\n\n--- ORANGE PAGE ---\n\n'.join(texts)[-40000:]),
    }
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding='utf-8')
    return path

def _canonical(src, account):
    q = _gb(src.get('quota_gb'), src.get('_quota_key','')) if isinstance(src, dict) else None
    u = _gb(src.get('used_gb'), src.get('_used_key','')) if isinstance(src, dict) else None
    r = _gb(src.get('remaining_gb'), src.get('_remaining_key','')) if isinstance(src, dict) else None
    pct = _num(src.get('usage_pct')) if isinstance(src, dict) else None
    if pct is None and q and u is not None: pct = u / q * 100
    if r is None and q is not None and u is not None: r = max(0, q-u)
    status = str(src.get('status') or 'Active').strip() if isinstance(src, dict) else 'Active'
    if pct is not None and pct >= 90 and status.lower() in ('active','normal','enabled',''):
        status = 'High Usage'
    due = str(src.get('due_date') or '').strip() if isinstance(src, dict) else ''
    if due and re.fullmatch(r'\d+(?:\.0+)?', due): due = f"{int(float(due))} Remaining Days"
    mobile = (account.get('login_username') or account.get('account_number') or '').strip()
    return {
        'provider': 'Orange',
        'site': account.get('site') or '',
        'account_number': account.get('account_number') or mobile,
        'landline': mobile,
        'service_type': 'Home Wireless',
        'package': str(src.get('package') or '').strip() if isinstance(src,dict) else '',
        'speed_mbps': 0,
        'quota_gb': round(q or 0, 2),
        'used_gb': round(u or 0, 2),
        'usage_pct': round(pct or 0, 2),
        'remaining_gb': round(r or 0, 2),
        'bill_egp': 0,
        'renewal_cost_egp': round(_num(src.get('renewal_cost_egp')) or 0, 2) if isinstance(src,dict) else 0,
        'outstanding_egp': round(_num(src.get('outstanding_egp')) or 0, 2) if isinstance(src,dict) else 0,
        'due_date': due,
        'renewal_date': str(src.get('renewal_date') or '').strip() if isinstance(src,dict) else '',
        'status': status,
    }

def _from_payloads(payloads, account):
    best = {}
    for payload in payloads:
        for obj in _walk(payload):
            for field in ALIASES:
                if field not in best or best[field] in (None,'',0,0.0):
                    v, k = _pick(obj, field)
                    if v not in (None,'',[],{}):
                        best[field] = v
                        if field in ('quota_gb','used_gb','remaining_gb'): best['_' + field.split('_')[0] + '_key'] = k
    return _canonical(best, account)

def _label_num(text, labels, unit=None):
    for label in labels:
        m = re.search(rf'{label}\s*[:\-]?\s*([0-9]+(?:[.,][0-9]+)?)\s*{unit or ""}', text, re.I)
        if m: return float(m.group(1).replace(',','.'))
    return None

def _parse_arabic_date(text):
    """Return DD-MM-YYYY from Orange Arabic renewal text when available."""
    months = {
        'يناير':1,'فبراير':2,'مارس':3,'أبريل':4,'ابريل':4,'مايو':5,'يونيو':6,
        'يوليو':7,'أغسطس':8,'اغسطس':8,'سبتمبر':9,'أكتوبر':10,'اكتوبر':10,'نوفمبر':11,'ديسمبر':12,
    }
    m = re.search(r'(\d{1,2})\s+(' + '|'.join(map(re.escape, months)) + r')\s+(\d{4})', text)
    if not m:
        return ''
    return f"{int(m.group(1)):02d}-{months[m.group(2)]:02d}-{m.group(3)}"


def _due_from_renewal_date(value):
    """Calculate Orange Due = Renewal Date - today's local date."""
    s = str(value or '').strip()
    if not s:
        return ''
    for fmt in ('%d-%m-%Y', '%d/%m/%Y', '%Y-%m-%d'):
        try:
            renewal = datetime.strptime(s, fmt).date()
            days = max(0, (renewal - date.today()).days)
            return f"{days} Remaining Days"
        except ValueError:
            pass
    return ''


def _from_text(text, account):
    """Orange Home Wireless mapping based on the visible My Orange Internet page.

    Source examples observed on orange.eg/ar/myaccount/internet:
      Home Wireless 350 GB – EGP 1,150
      السعه المتبقية:329559MBs
      تاريخ التجديد:1 أكتوبر 2026

    For Home Wireless, Remaining is converted MB -> GB using 1024 MB = 1 GB.
    Used is deliberately calculated as Quota - Remaining, per the dashboard rule.
    """
    out = {}

    # 1) Exact Home Wireless plan header: package, quota and renewal cost.
    plan = re.search(
        r'(Home\s+Wireless\s+([0-9]+(?:[.,][0-9]+)?)\s*GB)\s*[–—-]\s*(?:EGP\s*)?([0-9][0-9,]*(?:\.[0-9]+)?)',
        text, re.I
    )
    if plan:
        out['package'] = re.sub(r'\s+', ' ', plan.group(1)).strip()
        out['quota_gb'] = float(plan.group(2).replace(',', '.'))
        out['renewal_cost_egp'] = float(plan.group(3).replace(',', ''))

    # Fallback package/header patterns.
    if not out.get('package'):
        m = re.search(r'(Home\s+Wireless\s+[0-9]+(?:[.,][0-9]+)?\s*GB)', text, re.I)
        if m:
            out['package'] = re.sub(r'\s+', ' ', m.group(1)).strip()
            q = re.search(r'([0-9]+(?:[.,][0-9]+)?)\s*GB', out['package'], re.I)
            if q: out['quota_gb'] = float(q.group(1).replace(',', '.'))

    # 2) Remaining capacity shown by Orange in MBs.
    rem_mb = None
    for pat in (
        r'السع[هة]\s*المتبقي[هة]\s*[:：]?\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*MBs?',
        r'السع[هة]\s*المتبقي[هة]\s*[:：]?\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*ميجا',
        r'remaining\s*(?:capacity)?\s*[:：]?\s*([0-9][0-9,]*(?:\.[0-9]+)?)\s*MBs?',
    ):
        m = re.search(pat, text, re.I)
        if m:
            rem_mb = float(m.group(1).replace(',', ''))
            break
    if rem_mb is not None:
        out['remaining_gb'] = rem_mb / 1024.0

    # 3) Renewal cost fallback if header did not contain it.
    if out.get('renewal_cost_egp') is None:
        for pat in (
            r'Home\s+Wireless[^\n\r]{0,80}?EGP\s*([0-9][0-9,]*(?:\.[0-9]+)?)',
            r'(?:قيمة\s*التجديد|سعر\s*الباقة)\s*[:：]?\s*(?:EGP\s*)?([0-9][0-9,]*(?:\.[0-9]+)?)',
        ):
            m = re.search(pat, text, re.I)
            if m:
                out['renewal_cost_egp'] = float(m.group(1).replace(',', ''))
                break

    # 4) Renewal date from Arabic/English visible text.
    m = re.search(r'تاريخ\s*التجديد\s*[:：]?\s*([^\n\r]{4,40})', text, re.I)
    if m:
        raw_date = m.group(1).strip()
        out['renewal_date'] = _parse_arabic_date(raw_date) or raw_date
    if not out.get('renewal_date'):
        for pat in [r'(?:renewal|expiry|due)\s*date\s*[:\-]?\s*(\d{1,2}[\-/]\d{1,2}[\-/]\d{2,4})']:
            m = re.search(pat, text, re.I)
            if m: out['renewal_date'] = m.group(1); break

    # Due for Orange Home Wireless is derived from Renewal Date and today's date.
    if out.get('renewal_date'):
        out['due_date'] = _due_from_renewal_date(out['renewal_date'])

    # 5) Generic fallbacks only for fields still missing.
    if out.get('quota_gb') is None:
        out['quota_gb'] = _label_num(text, [r'total\s*(?:quota|package)',r'quota',r'allowance',r'سع[هة]\s*الباقة\s*الأساسية',r'اجمالي\s*الباقة',r'إجمالي\s*الباقة'], r'(?:GB|جيجا|جيجابايت)?')
    if out.get('remaining_gb') is None:
        out['remaining_gb'] = _label_num(text, [r'remaining',r'available',r'المتبقي|متبقي'], r'(?:GB|جيجا|جيجابايت)?')
    out['outstanding_egp'] = _label_num(text, [r'current\s*balance',r'outstanding',r'amount\s*due',r'balance',r'الرصيد|المستحق'], r'(?:EGP|LE|جنيه)?')

    # Orange Home Wireless business rule: Used = Quota - Remaining.
    q = _num(out.get('quota_gb'))
    r = _num(out.get('remaining_gb'))
    if q is not None and r is not None:
        out['used_gb'] = max(0.0, q - r)
        out['usage_pct'] = (out['used_gb'] / q * 100.0) if q else 0.0

    out['status'] = 'Active'
    return _canonical(out, account)

def _merge(a, b):
    out = dict(a)
    for k, v in b.items():
        if k in ('provider','site','account_number','landline','service_type'):
            if v: out[k] = v
        elif v not in (None,'',0,0.0): out[k] = v
    q = float(out.get('quota_gb') or 0); u = float(out.get('used_gb') or 0)
    if not out.get('usage_pct') and q: out['usage_pct'] = round(u/q*100, 2)
    if not out.get('remaining_gb') and q and u: out['remaining_gb'] = round(max(0,q-u), 2)
    if float(out.get('usage_pct') or 0) >= 90 and str(out.get('status') or '').lower() in ('active','normal','enabled',''):
        out['status'] = 'High Usage'
    return out

def _has_data(x):
    return bool((x.get('package') or '').strip() or float(x.get('quota_gb') or 0)>0 or float(x.get('remaining_gb') or 0)>0 or float(x.get('outstanding_egp') or 0)>0 or x.get('renewal_date'))

def _body(page):
    try: return page.locator('body').inner_text(timeout=10000)
    except Exception: return ''

def _is_login(page, text=''):
    low = (page.url + '\n' + text).lower()
    return ('/myaccount/login' in low or '/captcha-login' in low or ('ادخل رقم الموبايل' in text and 'كلمة السر' in text) or ('enter your mobile number' in low and 'password' in low))

def _fill_login(page, mobile, password):
    # Use broad selectors because Orange changes generated element names periodically.
    user = page.locator('input[type="tel"]:visible, input[inputmode="tel"]:visible, input[name*="mobile" i]:visible, input[id*="mobile" i]:visible, input[type="text"]:visible').first
    pw = page.locator('input[type="password"]:visible').first
    if not user.count() or not pw.count(): return False
    user.fill(mobile); pw.fill(password)
    btn = page.locator('button[type="submit"]:visible, input[type="submit"]:visible').first
    if btn.count(): btn.click()
    else: pw.press('Enter')
    return True

def _fetch_orange_lines_once(account, interactive=False):
    """Fetch Orange Egypt Home Wireless data through Orange My Account.

    Home Wireless is treated as a mobile/data SIM line. Login therefore uses the
    Orange mobile number + Orange account password, not the Home DSL portal.
    Saved browser storage is reused for scheduled background syncs.
    """
    mobile = (account.get('login_username') or account.get('account_number') or '').strip()
    password = (account.get('password') or '').strip()
    if not mobile: raise RuntimeError('Orange Home Wireless mobile number is missing.')
    if not password and not _session_path(account).exists():
        raise RuntimeError('Orange account password is required for the first Home Wireless login.')

    state = _session_path(account); payloads=[]; urls=[]; texts=[]
    with sync_playwright() as p:
        kwargs = {'headless': not interactive}
        if os.getenv('WE_USE_INSTALLED_CHROME','1').strip().lower() not in ('0','false','no'): kwargs['channel'] = 'chrome'
        try: browser = p.chromium.launch(**kwargs)
        except Exception:
            kwargs.pop('channel', None); browser = p.chromium.launch(**kwargs)
        ctx_kwargs = {'viewport': {'width':1440,'height':1000}, 'locale':'ar-EG'}
        if state.exists(): ctx_kwargs['storage_state'] = str(state)
        context = browser.new_context(**ctx_kwargs); page = context.new_page()

        def on_resp(resp):
            try:
                ctype = (resp.headers.get('content-type') or '').lower(); rt = resp.request.resource_type
                if rt not in ('xhr','fetch') and 'json' not in ctype: return
                urls.append({'url': resp.url.split('?')[0], 'status':resp.status, 'resource_type':rt, 'content_type':ctype})
                try: data = resp.json()
                except Exception:
                    try:
                        t = resp.text().lstrip(); data = json.loads(t) if t.startswith(('{','[')) else None
                    except Exception: data = None
                if isinstance(data,(dict,list)): payloads.append(data)
            except Exception: pass
        page.on('response', on_resp)

        try:
            # First try the authenticated account using a saved session.
            page.goto(ORANGE_ACCOUNT_URL, wait_until='domcontentloaded', timeout=60000); page.wait_for_timeout(1500)
            txt = _body(page); texts.append(txt)
            if _is_login(page, txt):
                if not password:
                    raise RuntimeError('Orange session expired. Use Edit Account to save the Orange password, then run Sync again.')
                if not _fill_login(page, mobile, password):
                    page.goto(ORANGE_LOGIN_URL, wait_until='domcontentloaded', timeout=60000); page.wait_for_timeout(1000)
                    if not _fill_login(page, mobile, password):
                        raise RuntimeError('Orange login form was not detected.')
                page.wait_for_timeout(2500)
                txt = _body(page); texts.append(txt)

                # Orange can route some accounts to a CAPTCHA login page.
                if _is_login(page, txt):
                    low = (page.url + '\n' + txt).lower()
                    captcha = ('captcha' in low or 'اكتب الكلمة بالأسفل' in txt or 'write the word below' in low)
                    if captcha and not interactive:
                        raise RuntimeError('Orange requires CAPTCHA verification. Run manual Sync for this line and complete the verification in Chrome.')
                    if captcha and interactive:
                        deadline = ORANGE_CAPTCHA_TIMEOUT_SECONDS * 1000
                        elapsed = 0
                        while elapsed < deadline:
                            page.wait_for_timeout(1500); elapsed += 1500
                            txt = _body(page)
                            if not _is_login(page, txt): break
                        texts.append(txt)
                    if _is_login(page, txt):
                        raise RuntimeError('Orange login was not completed. Check the mobile number/password or complete any verification shown in Chrome.')

            # Orange Home Wireless source of truth: the authenticated Internet page.
            # Do not mix unrelated account/tariff pages into the mapping because they can
            # contain balances and package names that belong to other mobile services.
            try:
                page.goto(ORANGE_INTERNET_URL, wait_until='domcontentloaded', timeout=60000)
                page.wait_for_timeout(1500)

                # Wait for the Home Wireless usage panel to render. Orange sometimes
                # loads this part asynchronously after the first DOMContentLoaded event.
                try:
                    page.wait_for_function(
                        """() => {
                            const t = (document.body && document.body.innerText) || '';
                            return /Home\s+Wireless/i.test(t) &&
                                   (/السع[هة]\s*المتبقي[هة]/.test(t) || /remaining/i.test(t));
                        }""",
                        timeout=30000
                    )
                except Exception:
                    # One reload helps when Orange returns the shell before account data.
                    page.reload(wait_until='domcontentloaded', timeout=60000)
                    page.wait_for_timeout(2500)

                internet_text = _body(page)
                texts.append(internet_text)

                # If Orange redirected to login from the Internet page, refresh the session
                # automatically in background using the stored credentials. This makes scheduled
                # Orange Home Wireless sync work without opening the dashboard, unless Orange
                # explicitly requires CAPTCHA/interactive verification.
                if _is_login(page, internet_text):
                    if not password:
                        raise RuntimeError('Orange session expired while opening the Internet page and no saved password is available.')
                    if not _fill_login(page, mobile, password):
                        page.goto(ORANGE_LOGIN_URL, wait_until='domcontentloaded', timeout=60000)
                        page.wait_for_timeout(1000)
                        if not _fill_login(page, mobile, password):
                            raise RuntimeError('Orange background login form was not detected.')
                    page.wait_for_timeout(2500)
                    login_text = _body(page); texts.append(login_text)
                    if _is_login(page, login_text):
                        low = (page.url + '\n' + login_text).lower()
                        captcha = ('captcha' in low or 'اكتب الكلمة بالأسفل' in login_text or 'write the word below' in low)
                        if captcha and not interactive:
                            raise RuntimeError('Orange requires CAPTCHA verification. Run manual Sync for this line once, then background sync will resume using the saved session.')
                        if captcha and interactive:
                            elapsed = 0; deadline = ORANGE_CAPTCHA_TIMEOUT_SECONDS * 1000
                            while elapsed < deadline:
                                page.wait_for_timeout(1500); elapsed += 1500
                                login_text = _body(page)
                                if not _is_login(page, login_text): break
                            texts.append(login_text)
                        if _is_login(page, login_text):
                            raise RuntimeError('Orange login was not completed while refreshing the background session.')
                    page.goto(ORANGE_INTERNET_URL, wait_until='domcontentloaded', timeout=60000)
                    page.wait_for_timeout(1800)
                    internet_text = _body(page); texts.append(internet_text)
                    if _is_login(page, internet_text):
                        raise RuntimeError('Orange Internet page still requires login after background session refresh.')

                # Require the page that actually contains the Home Wireless package.
                if not re.search(r'Home\s+Wireless', internet_text, re.I):
                    diag = _diag(account, urls, payloads, texts)
                    raise RuntimeError(f'Orange Internet page opened but Home Wireless details were not found. Diagnostic saved to {diag}.')

            except RuntimeError:
                raise
            except Exception as e:
                diag = _diag(account, urls, payloads, texts)
                raise RuntimeError(f'Could not read Orange Home Wireless Internet page: {e}. Diagnostic saved to {diag}.')

            try: context.storage_state(path=str(state))
            except Exception: pass

            # DOM text from /ar/myaccount/internet has priority over API payloads.
            # API payloads are only a fallback for fields not present in the visible page.
            dom_data = _from_text(internet_text, account)
            api_data = _from_payloads(payloads, account)
            line_data = _merge(api_data, dom_data)

            # Home Wireless business rule requested for Orange only:
            # Used = Quota - Remaining, where Remaining is converted MB -> GB /1024.
            q = float(line_data.get('quota_gb') or 0)
            r = float(line_data.get('remaining_gb') or 0)
            if q > 0 and r >= 0:
                used = max(0.0, q - r)
                line_data['used_gb'] = round(used, 2)
                line_data['usage_pct'] = round((used / q * 100.0) if q else 0.0, 2)

            if line_data.get('renewal_date'):
                line_data['due_date'] = _due_from_renewal_date(line_data['renewal_date'])

            if not _has_data(line_data) or not line_data.get('package') or not line_data.get('quota_gb'):
                diag = _diag(account, urls, payloads, texts)
                raise RuntimeError(f'Orange Home Wireless Internet page was reached, but required fields were not mapped. Diagnostic saved to {diag}.')
            return [line_data]
        finally:
            try: context.close(); browser.close()
            except Exception: pass


def fetch_orange_lines(account, interactive=False):
    """Sync Orange silently first; open Chrome only when Orange actually requires CAPTCHA.

    Manual Sync no longer implies a visible browser. Both manual and scheduled syncs first
    run headless. A visible Chrome window is launched only after the headless attempt
    positively detects Orange CAPTCHA/verification.
    """
    try:
        return _fetch_orange_lines_once(account, interactive=False)
    except RuntimeError as exc:
        msg = str(exc)
        captcha_required = (
            'requires CAPTCHA verification' in msg
            or 'CAPTCHA' in msg and 'verification' in msg
        )
        if interactive and captcha_required:
            return _fetch_orange_lines_once(account, interactive=True)
        raise
