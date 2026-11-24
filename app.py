import os
import sqlite3
import threading
import time
import shutil
import io
from pathlib import Path
from datetime import datetime
from flask import Flask, jsonify, render_template_string, request, send_file
from dotenv import load_dotenv
from cryptography.fernet import Fernet, InvalidToken
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

load_dotenv()
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def _resolve_data_dir():
    configured = (os.getenv('WE_DATA_DIR') or '').strip()
    if configured:
        return os.path.abspath(os.path.expandvars(os.path.expanduser(configured)))
    if os.name == 'nt':
        root = os.path.join(os.path.expanduser('~'), 'WE-Dashboard-Data')
    else:
        root = os.path.join(os.path.expanduser('~'), '.we-dashboard-data')
    return os.path.abspath(root)

DATA_DIR = _resolve_data_dir()
os.makedirs(DATA_DIR, exist_ok=True)
BACKUP_DIR = os.path.join(DATA_DIR, 'backups')
SESSIONS_DIR = os.path.join(DATA_DIR, 'sessions')
DIAGNOSTICS_DIR = os.path.join(DATA_DIR, 'diagnostics')
for _folder in (BACKUP_DIR, SESSIONS_DIR, DIAGNOSTICS_DIR):
    os.makedirs(_folder, exist_ok=True)
DB = os.path.join(DATA_DIR, 'we_dashboard.db')
KEY_FILE = os.path.join(DATA_DIR, '.we_secret.key')

def migrate_legacy_local_data():
    """One-time migration from an older version folder into the persistent data folder."""
    legacy_db = os.path.join(BASE_DIR, 'we_dashboard.db')
    legacy_key = os.path.join(BASE_DIR, '.we_secret.key')
    legacy_sessions = os.path.join(BASE_DIR, 'sessions')
    legacy_diagnostics = os.path.join(BASE_DIR, 'diagnostics')
    try:
        if not os.path.exists(DB) and os.path.exists(legacy_db):
            shutil.copy2(legacy_db, DB)
            print('Migrated legacy database to:', DB)
        if not os.path.exists(KEY_FILE) and os.path.exists(legacy_key):
            shutil.copy2(legacy_key, KEY_FILE)
            print('Migrated legacy encryption key to:', KEY_FILE)
        if os.path.isdir(legacy_sessions) and not any(Path(SESSIONS_DIR).iterdir()):
            shutil.copytree(legacy_sessions, SESSIONS_DIR, dirs_exist_ok=True)
            print('Migrated saved WE sessions to:', SESSIONS_DIR)
        if os.path.isdir(legacy_diagnostics) and not any(Path(DIAGNOSTICS_DIR).iterdir()):
            shutil.copytree(legacy_diagnostics, DIAGNOSTICS_DIR, dirs_exist_ok=True)
    except Exception as e:
        print('Legacy data migration warning:', e)

migrate_legacy_local_data()
app = Flask(__name__)
APP_BUILD = 'Internet Dashboard v5.33 - Internet & Land Lines'
SYNC_LOCK = threading.Lock()
SYNC_STATE = {'running': False, 'last_started': None, 'last_finished': None, 'last_error': '', 'updated': 0, 'failed': 0}

HTML = r'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Internet Dashboard</title>
<style>
:root{--we:#5b2d90;--we2:#7c4bb2;--bg:#f4f6fb;--card:#fff;--text:#18212f;--muted:#70798a;--line:#e8ebf2;--ok:#16784b;--warn:#a76500;--bad:#b42318;--blue:#1769aa}
*{box-sizing:border-box}body{font-family:Segoe UI,Arial,sans-serif;background:var(--bg);margin:0;color:var(--text)}.wrap{max-width:1650px;margin:auto;padding:22px}.top{display:flex;gap:14px;justify-content:space-between;align-items:center;margin-bottom:16px}.title{font-size:27px;font-weight:750}.sub{color:var(--muted);font-size:13px;margin-top:4px}.actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.btn{border:0;border-radius:9px;padding:10px 14px;background:var(--we);color:#fff;cursor:pointer;font-weight:600}.btn.secondary{background:#eef0f6;color:#333}.btn.outline{background:#fff;color:var(--we);border:1px solid #d9c9e9}.btn.danger{background:#b42318}.btn.small{padding:7px 9px;font-size:11px}.sync{font-size:12px;color:var(--muted)}
.filters{display:grid;grid-template-columns:1.5fr 1fr 1fr 1fr 1fr 1fr auto;gap:10px;background:#fff;padding:12px;border-radius:13px;box-shadow:0 2px 10px #0000000a;margin-bottom:14px}.filters input,.filters select,.form input,.form select{border:1px solid #dce1ea;border-radius:8px;padding:10px;background:white;min-width:0;width:100%}.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:12px}.card{background:var(--card);border-radius:13px;padding:16px;box-shadow:0 2px 10px #0000000a}.k{font-size:12px;color:var(--muted)}.v{font-size:24px;font-weight:750;margin-top:7px}.mini{font-size:11px;color:var(--muted);margin-top:4px}.layout{display:grid;grid-template-columns:minmax(0,2.7fr) minmax(320px,1fr);gap:14px;margin-top:14px}.layout.detail-hidden{grid-template-columns:1fr}.detail-panel.hidden{display:none}.panel-title-row{display:flex;align-items:center;justify-content:space-between;gap:10px}.panel-title-row h3{margin:0}.toggle-detail-btn{white-space:nowrap}.panel{background:#fff;border-radius:13px;padding:16px;box-shadow:0 2px 10px #0000000a}.panel h3{font-size:15px;margin:0 0 10px}.tablewrap{overflow:auto;max-height:620px}.tbl{width:100%;border-collapse:collapse;font-size:12.5px;white-space:nowrap}.tbl th,.tbl td{padding:10px 8px;border-bottom:1px solid var(--line);text-align:left}.tbl th{position:sticky;top:0;background:#fff;color:#667085;z-index:2}.tbl tr{cursor:pointer}.tbl tr:hover{background:#faf8ff}.pill{padding:4px 8px;border-radius:20px;font-size:11px;font-weight:650;display:inline-block}.ok{background:#def7ea;color:var(--ok)}.warn{background:#fff0cf;color:var(--warn)}.bad{background:#fde2df;color:var(--bad)}.info{background:#e6f0fb;color:var(--blue)}.usagebar{width:100px;height:7px;background:#e8eaf0;border-radius:7px;overflow:hidden;display:inline-block;vertical-align:middle;margin-right:5px}.usagebar i{display:block;height:100%;background:var(--we)}
.detail-empty{color:var(--muted);font-size:13px;padding:35px 5px;text-align:center}.dhead{font-weight:750;font-size:18px}.dline{color:var(--muted);font-size:12px;margin:4px 0 15px}.dgrid{display:grid;grid-template-columns:1fr 1fr;gap:8px}.ditem{background:#f7f8fb;border-radius:9px;padding:10px}.ditem b{display:block;font-size:11px;color:var(--muted);font-weight:500}.ditem span{display:block;margin-top:5px;font-size:13px;font-weight:650}.section{font-size:11px;text-transform:uppercase;letter-spacing:.06em;color:#8b93a3;margin:18px 0 8px}.account{display:flex;justify-content:space-between;gap:8px;border-bottom:1px solid var(--line);padding:10px 0;font-size:12px;align-items:center}.account:last-child{border:0}.account small{color:var(--muted)}.account-actions{display:flex;gap:5px}.row-actions{display:flex;gap:5px}.source{font-size:10px;color:#777}
.modal{position:fixed;inset:0;background:#11182788;display:none;align-items:center;justify-content:center;padding:18px;z-index:100}.modal.open{display:flex}.modal-card{background:#fff;border-radius:14px;box-shadow:0 20px 70px #0004;width:min(720px,100%);max-height:92vh;overflow:auto}.modal-head{display:flex;justify-content:space-between;align-items:center;padding:16px 18px;border-bottom:1px solid var(--line)}.modal-head h2{font-size:18px;margin:0}.x{border:0;background:transparent;font-size:25px;cursor:pointer;color:#777}.form{padding:18px}.form-grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}.field label{display:block;font-size:11px;color:var(--muted);margin:0 0 5px}.span2{grid-column:1/-1}.form-foot{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}.hint{font-size:11px;color:var(--muted);line-height:1.5;margin-top:5px}.toast{position:fixed;right:20px;bottom:20px;background:#111827;color:#fff;padding:11px 15px;border-radius:9px;box-shadow:0 5px 25px #0003;display:none;z-index:200}.toast.show{display:block}
.alert-panel{background:#fff;border-radius:13px;padding:14px 16px;box-shadow:0 2px 10px #0000000a;margin-bottom:14px;display:none}.alert-panel.show{display:block}.alert-head{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:10px}.alert-title{font-weight:750}.alert-actions{display:flex;gap:7px;flex-wrap:wrap}.alert-list{display:grid;gap:8px}.alert-item{border-radius:10px;padding:10px 12px;border-left:4px solid #f59e0b;background:#fffaf0;font-size:12px}.alert-item.renew{border-left-color:#b42318;background:#fff6f5}.alert-item b{display:block;font-size:13px;margin-bottom:2px}.badge-count{background:#b42318;color:#fff;border-radius:999px;padding:2px 7px;font-size:11px}.btn.tiny{padding:6px 9px;font-size:11px}

.notif-wrap,.settings-wrap{position:relative;display:inline-flex;align-items:center}.icon-btn{position:relative;width:40px;height:40px;border-radius:10px;border:1px solid #d9c9e9;background:#fff;color:var(--we);cursor:pointer;font-size:19px;display:inline-flex;align-items:center;justify-content:center}.notif-badge{position:absolute;top:-6px;right:-6px;min-width:19px;height:19px;padding:0 5px;border-radius:999px;background:#d92d20;color:#fff;font-size:10px;font-weight:800;display:none;align-items:center;justify-content:center;border:2px solid #fff}.notif-badge.show{display:flex}.dropdown{position:absolute;right:0;top:48px;width:min(420px,90vw);background:#fff;border:1px solid #e4e7ec;border-radius:13px;box-shadow:0 18px 50px #10182826;z-index:500;display:none;overflow:hidden}.dropdown.open{display:block}.dropdown-head{display:flex;justify-content:space-between;align-items:center;padding:12px 14px;border-bottom:1px solid #eef0f4;font-weight:750}.dropdown-list{max-height:420px;overflow:auto}.notif-row{padding:12px 14px;border-bottom:1px solid #f0f1f4;cursor:pointer;background:#fff}.notif-row.unread{background:#faf7ff}.notif-row:hover{background:#f7f4fb}.notif-row-title{font-size:13px;font-weight:800;color:#000}.notif-row-msg{font-size:11px;color:#000;font-weight:700;margin-top:4px;line-height:1.45}.notif-row-meta{font-size:10px;color:#000;font-weight:700;margin-top:5px}.notif-dot{display:inline-block;width:7px;height:7px;border-radius:50%;background:#7c4bb2;margin-right:6px}.notif-row:not(.unread) .notif-dot{background:#d0d5dd}.empty-dd{padding:25px 14px;text-align:center;color:#98a2b3;font-size:12px}.link-btn{border:0;background:none;color:var(--we);font-size:11px;cursor:pointer;font-weight:650}.settings-menu{width:220px}.settings-menu a,.settings-menu button{display:block;width:100%;padding:11px 13px;text-align:left;border:0;background:#fff;color:#344054;text-decoration:none;cursor:pointer;font-size:12px}.settings-menu a:hover,.settings-menu button:hover{background:#f7f4fb}.settings-btn{white-space:nowrap}.caret{font-size:10px;margin-left:4px}

/* Unified application shell matching Visual Dashboard */
:root{--nav:#123d69;--nav2:#0b2d50;--orange:#ff7900;--shell-blue:#174f83;--shell-ink:#16345c}
.app-shell{display:grid;grid-template-columns:195px 1fr;min-height:100vh}.side-shell{background:linear-gradient(180deg,var(--nav),var(--nav2));color:#fff;padding:18px 0;position:sticky;top:0;height:100vh;z-index:20}.side-brand{padding:0 18px 20px;font-weight:800;font-size:22px;line-height:1.05}.side-brand small{display:block;font-size:12px;font-weight:600;opacity:.9;margin-top:5px}.side-nav a{display:block;padding:15px 18px;color:#eaf3ff;text-decoration:none;border-left:4px solid transparent;font-size:14px}.side-nav a.active,.side-nav a:hover{background:#ffffff14;border-left-color:#ff9c40}.side-ver{position:absolute;bottom:12px;left:14px;font-size:10px;opacity:.75}.main-shell{min-width:0}.main-shell>.wrap{max-width:none;margin:0;padding:0 20px 30px}.main-shell .top{min-height:78px;background:linear-gradient(90deg,var(--shell-blue),var(--nav));margin:0 -20px 20px;padding:13px 24px;color:#fff}.main-shell .title{color:#fff}.main-shell .sub,.main-shell .sync{color:#e4edf7}.main-shell .top .btn.outline,.main-shell .top .btn.secondary,.main-shell .top .icon-btn{background:#fff;color:var(--shell-ink);border-color:#d5e1ec}.panel,.card,.filters,.alert-panel{border:1px solid #dce5ef;box-shadow:0 4px 18px #0f27430c}.main-shell .btn:not(.secondary):not(.outline):not(.danger){background:var(--shell-blue)}
.tbl th.sortable{cursor:pointer;user-select:none}.tbl th.sortable:hover{color:#174f83;background:#f4f8fc}.tbl th.sortable span{font-size:10px;margin-left:3px;color:#174f83}
@media(max-width:1100px){.app-shell{grid-template-columns:1fr}.side-shell{display:none}.main-shell>.wrap{padding:0 12px 22px}.main-shell .top{margin:0 -12px 15px}}
@media(max-width:1200px){.cards{grid-template-columns:repeat(3,1fr)}.layout{grid-template-columns:1fr}.filters{grid-template-columns:1fr 1fr 1fr}}@media(max-width:700px){.wrap{padding:12px}.top{align-items:flex-start;flex-direction:column}.cards{grid-template-columns:repeat(2,1fr)}.filters{grid-template-columns:1fr}.dgrid,.form-grid{grid-template-columns:1fr}.span2{grid-column:auto}}
</style></head><body><div class="app-shell"><aside class="side-shell"><div class="side-brand">Internet Dashboard<small>WE & Orange</small></div><nav class="side-nav"><a class="active" href="/">Internet Lines</a><a href="/landlines">Land Lines</a><a href="/visual">Visual Dashboard</a><a href="/reports">Reports</a></nav><div class="side-ver">Unified Interface v5.33</div></aside><main class="main-shell"><div class="wrap">
<div class="top"><div><div class="title">Internet Dashboard</div><div class="sub">Internet lines operational view</div></div><div class="actions"><span class="sync" id="sync"></span><button class="btn outline" onclick="openAccount()">+ Add Internet Account</button><button class="btn" id="syncBtn" onclick="syncNow()">Sync</button><button class="btn outline" onclick="exportLines('xlsx')">Export Excel</button><div class="notif-wrap"><button class="icon-btn" title="Notifications" onclick="toggleNotifications(event)">🔔<span id="notifBadge" class="notif-badge"></span></button><div id="notifDropdown" class="dropdown"><div class="dropdown-head"><span>Notifications</span><button class="link-btn" onclick="markAllRead(event)">Mark all read</button></div><div id="notifList" class="dropdown-list"></div></div></div><button class="btn secondary" id="refreshBtn" onclick="loadData(true)">Refresh View</button><button class="btn danger" id="stopBgBtn" onclick="stopBackground()">Stop Background</button></div></div>

<div class="filters"><input id="q" placeholder="Search provider, site, legal entity, branch, landline..." oninput="render()"><select id="site" onchange="render()"><option value="">All Sites</option></select><select id="legal_entity" onchange="render()"><option value="">All Legal Entities</option></select><select id="account_name" onchange="render()"><option value="">All Branchs</option></select><select id="provider" onchange="render()"><option value="">All Providers</option><option value="WE">WE</option><option value="Orange">Orange Home Wireless</option></select><select id="status" onchange="render()"><option value="">All Status</option><option>Active</option><option>High Usage</option><option>Suspended</option><option>Overdue</option></select><button class="btn secondary" onclick="clearFilters()">Clear</button></div>
<div id="cards" class="cards"></div>
<div class="layout" id="mainLayout"><div class="panel"><div class="panel-title-row"><h3>All Lines</h3><button class="btn small secondary toggle-detail-btn" id="showDetailBtn" onclick="setDetailPanel(true)" style="display:none">Show Selected Line</button></div><div class="tablewrap"><table class="tbl"><thead><tr><th class="sortable" onclick="sortBy('legal_entity')">Legal Entity <span data-sort="legal_entity"></span></th><th class="sortable" onclick="sortBy('site')">Site <span data-sort="site"></span></th><th class="sortable" onclick="sortBy('account_name')">Branch <span data-sort="account_name"></span></th><th class="sortable" onclick="sortBy('landline')">Landline <span data-sort="landline"></span></th><th class="sortable" onclick="sortBy('provider')">Provider <span data-sort="provider"></span></th><th class="sortable" onclick="sortBy('service_type')">Service <span data-sort="service_type"></span></th><th class="sortable" onclick="sortBy('package')">Package <span data-sort="package"></span></th><th class="sortable" onclick="sortBy('quota_gb')">Quota <span data-sort="quota_gb"></span></th><th class="sortable" onclick="sortBy('usage_pct')">Usage <span data-sort="usage_pct"></span></th><th class="sortable" onclick="sortBy('remaining_gb')">Remaining <span data-sort="remaining_gb"></span></th><th class="sortable" onclick="sortBy('renewal_cost_egp')">Renewal Cost <span data-sort="renewal_cost_egp"></span></th><th class="sortable" onclick="sortBy('due_date')">Due <span data-sort="due_date"></span></th><th class="sortable" onclick="sortBy('last_sync')">Last Sync <span data-sort="last_sync"></span></th><th class="sortable" onclick="sortBy('status')">Status <span data-sort="status"></span></th><th>Manage</th></tr></thead><tbody id="rows"></tbody></table></div></div><div class="panel detail-panel" id="detailPanel"><div class="panel-title-row"><h3>Selected Line</h3><button class="btn small secondary toggle-detail-btn" onclick="setDetailPanel(false)">Hide</button></div><div id="detail" class="detail-empty">Select any line to show its full details here.</div><div class="section">Internet Accounts</div><div id="accounts"></div></div></div>
</div>
<div class="modal" id="lineModal"><div class="modal-card"><div class="modal-head"><h2 id="lineTitle">Add Line</h2><button class="x" onclick="closeModal('lineModal')">×</button></div><form class="form" id="lineForm" onsubmit="saveLine(event)"><input type="hidden" id="line_id"><div class="form-grid">
<div class="field"><label>Site / Branch *</label><input id="line_site" required placeholder="Head Office"></div><div class="field"><label>Legal Entity</label><input id="line_legal_entity" placeholder="Linah Farms"></div><input type="hidden" id="line_account"><datalist id="accountList"></datalist>
<div class="field"><label>Landline *</label><input id="line_landline" required placeholder="02-xxxx-xxxx"></div><div class="field"><label>Service Type</label><select id="line_service"><option>Internet + Landline</option><option>Internet</option><option>Internet Backup</option><option>Landline</option></select></div>
<div class="field"><label>Package</label><input id="line_package" placeholder="WE Space 250"></div><div class="field"><label>Status</label><select id="line_status"><option>Active</option><option>High Usage</option><option>Suspended</option><option>Overdue</option></select></div>
<div class="field"><label>Speed (Mbps)</label><input id="line_speed" type="number" min="0" step="1" value="0"></div><div class="field"><label>Quota (GB)</label><input id="line_quota" type="number" min="0" step="0.01" value="0"></div>
<div class="field"><label>Used (GB)</label><input id="line_used" type="number" min="0" step="0.01" value="0"></div><div class="field"><label>Renewal Cost (EGP)</label><input id="line_renewal_cost" type="number" min="0" step="0.01" value="0"></div>
<div class="field"><label>Outstanding / Current Balance (EGP)</label><input id="line_outstanding" type="number" min="0" step="0.01" value="0"></div><div class="field"><label>Due</label><input id="line_due" type="text" placeholder="9 Remaining Days"></div>
<div class="field"><label>Renewal Date</label><input id="line_renewal" type="date"></div>
</div><div class="form-foot"><button type="button" class="btn secondary" onclick="closeModal('lineModal')">Cancel</button><button class="btn" type="submit">Save Line</button></div></form></div></div>
<div class="modal" id="accountModal"><div class="modal-card"><div class="modal-head"><h2 id="accountTitle">Add Internet Account</h2><button class="x" onclick="closeModal('accountModal')">×</button></div><form class="form" id="accountForm" onsubmit="saveAccount(event)"><input type="hidden" id="account_id"><div class="form-grid">
<div class="field"><label>Provider *</label><select id="acc_provider" onchange="updateAccountProviderUI()"><option value="WE" selected>WE</option><option value="Orange">Orange Home Wireless</option></select></div><div class="field"><label>Branch</label><input id="acc_name" placeholder="Head Office"></div><input type="hidden" id="acc_number">
<div class="field"><label>Default Site / Branch</label><input id="acc_site" placeholder="Head Office"></div><div class="field"><label>Default Legal Entity</label><input id="acc_legal_entity" placeholder="Linah Farms"><div class="hint">Applied automatically to lines created from this internet account.</div></div><div class="field"><label id="acc_login_label">WE Service Number (Login) *</label><input id="acc_username" autocomplete="off" required placeholder="0237487976"></div>
<div class="field" id="serviceTypeField"><label>Portal Service Type</label><select id="acc_portal_type"><option value="Internet" selected>Internet</option><option value="LandLine">LandLine</option><option value="Home Wireless">Home Wireless</option></select><div class="hint" id="acc_portal_hint">WE: choose Internet. Orange Home Wireless is treated as a mobile/data SIM line.</div></div>
<div class="field span2"><label id="acc_password_label">Portal Password</label><input id="acc_password" type="password" autocomplete="new-password" placeholder="Leave blank when editing to keep existing password"><div class="hint" id="acc_password_hint">Stored locally in encrypted form. Orange Home Wireless uses the Orange mobile account password.</div></div>
<div class="field"><label>Automatic Sync</label><select id="acc_enabled"><option value="1">Enabled</option><option value="0">Disabled</option></select></div>
</div><div class="form-foot"><button type="button" class="btn secondary" onclick="closeModal('accountModal')">Cancel</button><button class="btn" type="submit">Save Account</button></div></form></div></div>
<div class="toast" id="toast"></div>
</div></main></div>
<script>
let ALL=[],DATA=null,SORT_KEY='',SORT_DIR=1;
const $=id=>document.getElementById(id), money=n=>(Number(n||0)).toLocaleString(undefined,{maximumFractionDigits:2});

function sortValue(x,k){
  const numeric=['quota_gb','usage_pct','remaining_gb','renewal_cost_egp'];
  if(numeric.includes(k)) return Number(x[k]||0);
  if(k==='due_date'){const m=String(x[k]||'').match(/-?\d+/);return m?Number(m[0]):999999;}
  return String(x[k]??'').toLowerCase();
}
function sortBy(k){SORT_DIR=(SORT_KEY===k)?-SORT_DIR:1;SORT_KEY=k;render()}
function sortedLines(lines){if(!SORT_KEY)return lines;return [...lines].sort((a,b)=>{const av=sortValue(a,SORT_KEY),bv=sortValue(b,SORT_KEY);if(typeof av==='number'&&typeof bv==='number')return (av-bv)*SORT_DIR;return String(av).localeCompare(String(bv),undefined,{numeric:true,sensitivity:'base'})*SORT_DIR})}
function renderSortIndicators(){document.querySelectorAll('[data-sort]').forEach(e=>e.textContent=e.dataset.sort===SORT_KEY?(SORT_DIR===1?'▲':'▼'):'↕')}

function esc(v){return String(v??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}
function pcls(s){if(s==='Active')return'ok';if(s==='High Usage')return'warn';if(s==='Suspended'||s==='Overdue')return'bad';return'info'}
function toast(m){$('toast').textContent=m;$('toast').classList.add('show');setTimeout(()=>$('toast').classList.remove('show'),2600)}
function openModal(id){$(id).classList.add('open')} function closeModal(id){$(id).classList.remove('open')}
function fillSelect(id,vals){const el=$(id),cur=el.value,first=el.options[0].outerHTML;el.innerHTML=first+[...new Set(vals.filter(Boolean))].sort().map(v=>`<option>${esc(v)}</option>`).join('');el.value=cur}
function exportLines(fmt){const p=new URLSearchParams();const ids=['q','site','legal_entity','account_name','provider','status'];ids.forEach(id=>{const e=$(id);if(e&&e.value)p.set(id,e.value)});window.location.href='/export/lines.'+fmt+(p.toString()?'?'+p.toString():'')}
function clearFilters(){['q','site','legal_entity','account_name','provider','status'].forEach(id=>$(id).value='');render()}
function renderNotifications(){const n=(DATA&&DATA.notifications)||[];const unread=(DATA&&Number(DATA.unread_notifications||0))||0;const badge=$('notifBadge'),list=$('notifList');if(badge){badge.textContent=unread>99?'99+':unread;badge.classList.toggle('show',unread>0)}if(list){list.innerHTML=n.length?n.map(a=>`<div class="notif-row ${a.read?'':'unread'}" onclick="markRead('${encodeURIComponent(a.key)}',event)"><div class="notif-row-title"><span class="notif-dot"></span>${esc(a.title)} — ${esc(a.landline)}</div><div class="notif-row-msg">${esc(a.message)}</div><div class="notif-row-meta">${esc(a.site||'-')} · ${a.read?'Read':'Unread'}</div></div>`).join(''):'<div class="empty-dd">No active notifications.</div>'}pushBrowserAlerts(n.filter(a=>!a.read))}
function closeMenus(){const n=$('notifDropdown');if(n)n.classList.remove('open')}function toggleNotifications(e){e.stopPropagation();const n=$('notifDropdown');if(n)n.classList.toggle('open')}document.addEventListener('click',closeMenus);document.addEventListener('keydown',e=>{if(e.key==='Escape')closeMenus()});
async function markRead(key,e){if(e)e.stopPropagation();try{const r=await fetch('/api/notifications/read',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({key:decodeURIComponent(key)})});if(r.ok)await loadData()}catch(_){}}async function markAllRead(e){if(e)e.stopPropagation();try{const r=await fetch('/api/notifications/read',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({all:true})});if(r.ok)await loadData()}catch(_){}}
async function enableBrowserAlerts(){if(!('Notification' in window))return toast('Browser notifications are not supported');const r=await Notification.requestPermission();toast(r==='granted'?'Browser alerts enabled':'Browser alerts not enabled')}
function pushBrowserAlerts(items){if(!('Notification' in window)||Notification.permission!=='granted')return;items.forEach(a=>{const k='wealert:'+a.key;if(localStorage.getItem(k)==='1')return;new Notification(`${a.title} - ${a.landline}`,{body:a.message});localStorage.setItem(k,'1')})}
function filtered(){const q=$('q').value.toLowerCase(),s=$('site').value,le=$('legal_entity').value,an=$('account_name').value,pv=$('provider').value,st=$('status').value;return ALL.filter(x=>(!q||[x.site,x.legal_entity,x.account_name,x.landline,x.account_number,x.package,x.service_type,x.provider].join(' ').toLowerCase().includes(q))&&(!s||x.site===s)&&(!le||x.legal_entity===le)&&(!an||x.account_name===an)&&(!pv||x.provider===pv)&&(!st||x.status===st))}
function setDetailPanel(show){const layout=$('mainLayout'),panel=$('detailPanel'),btn=$('showDetailBtn');if(!layout||!panel||!btn)return;panel.classList.toggle('hidden',!show);layout.classList.toggle('detail-hidden',!show);btn.style.display=show?'none':'inline-flex';localStorage.setItem('we_detail_panel',show?'show':'hide')}function initDetailPanel(){setDetailPanel(localStorage.getItem('we_detail_panel')!=='hide')}
function showDetail(id){const x=ALL.find(y=>y.id===id);if(!x)return;$('detail').innerHTML=`<div class="dhead">${esc(x.site)}</div><div class="dline">${esc(x.landline)}</div><div class="dgrid"><div class="ditem"><b>Legal Entity</b><span>${esc(x.legal_entity||'-')}</span></div><div class="ditem"><b>Branch</b><span>${esc(x.account_name||'-')}</span></div><div class="ditem"><b>Provider</b><span>${esc(x.provider||'WE')}</span></div><div class="ditem"><b>Service</b><span>${esc(x.service_type)}</span></div><div class="ditem"><b>Status</b><span><span class="pill ${pcls(x.status)}">${esc(x.status)}</span></span></div><div class="ditem"><b>Package</b><span>${esc(x.package)}</span></div><div class="ditem"><b>Speed</b><span>${x.speed_mbps||0} Mbps</span></div><div class="ditem"><b>Quota</b><span>${x.quota_gb||0} GB</span></div><div class="ditem"><b>Usage</b><span>${Number(x.usage_pct||0).toFixed(1)}% (${x.used_gb||0} GB)</span></div><div class="ditem"><b>Remaining</b><span>${x.remaining_gb||0} GB</span></div><div class="ditem"><b>Renewal Cost</b><span>${money(x.renewal_cost_egp)} EGP</span></div><div class="ditem"><b>Outstanding</b><span>${money(x.outstanding_egp)} EGP</span></div><div class="ditem"><b>Due</b><span>${esc(x.due_date||'-')}</span></div><div class="ditem"><b>Renewal Date</b><span>${esc(x.renewal_date||'-')}</span></div><div class="ditem"><b>Source</b><span>${esc(x.source||'manual')}</span></div><div class="ditem"><b>Last Sync</b><span>${esc(x.last_sync||'-')}</span></div></div>`}
function render(){const lines=sortedLines(filtered()),active=lines.filter(x=>x.status==='Active').length,high=lines.filter(x=>Number(x.usage_pct)>=90).length;$('cards').innerHTML=`<div class=card><div class=k>Visible Lines</div><div class=v>${lines.length}</div><div class=mini>of ${ALL.length} total</div></div><div class=card><div class=k>Active</div><div class=v>${active}</div></div><div class=card><div class=k>High Usage</div><div class=v>${high}</div><div class=mini>90% or more</div></div>`;
$('rows').innerHTML=lines.map(x=>`<tr onclick="showDetail(${x.id})"><td>${esc(x.legal_entity||'-')}</td><td>${esc(x.site||'-')}</td><td>${esc(x.account_name||'-')}</td><td>${esc(x.landline)}</td><td>${esc(x.provider||'WE')}</td><td>${esc(x.service_type)}</td><td>${esc(x.package)}</td><td>${Number(x.quota_gb||0).toLocaleString(undefined,{maximumFractionDigits:2})} GB</td><td><span class="usagebar"><i style="width:${Math.min(100,Number(x.usage_pct||0))}%"></i></span>${Number(x.usage_pct||0).toFixed(1)}%</td><td>${x.remaining_gb||0} GB</td><td>${money(x.renewal_cost_egp)} EGP</td><td>${esc(x.due_date||'-')}</td><td>${esc(x.last_sync||'-')}</td><td><span class="pill ${pcls(x.status)}">${esc(x.status)}</span></td><td><div class="row-actions"><button class="btn small" id="lineSyncBtn-${x.id}" onclick="event.stopPropagation();syncLine(${x.id})">Sync</button><button class="btn small secondary" onclick="event.stopPropagation();editLine(${x.id})">Edit</button><button class="btn small danger" onclick="event.stopPropagation();deleteLine(${x.id})">Delete</button></div></td></tr>`).join('')||'<tr><td colspan="15">No lines match the selected filters.</td></tr>';renderSortIndicators()}
function openLine(x=null){$('lineForm').reset();$('line_id').value=x?.id||'';$('lineTitle').textContent=x?'Edit Line':'Add Line';$('line_service').value=x?.service_type||'Internet + Landline';$('line_status').value=x?.status||'Active';['site','legal_entity','account_number','landline','package','speed_mbps','quota_gb','used_gb','renewal_cost_egp','outstanding_egp','due_date','renewal_date'].forEach(k=>{const m={account_number:'account',speed_mbps:'speed',quota_gb:'quota',used_gb:'used',renewal_cost_egp:'renewal_cost',outstanding_egp:'outstanding',due_date:'due',renewal_date:'renewal'};const id='line_'+(m[k]||k);if($(id))$(id).value=x?.[k]??($(id).type==='number'?0:'')});openModal('lineModal')}
function editLine(id){openLine(ALL.find(x=>x.id===id))}
async function saveLine(e){e.preventDefault();const id=$('line_id').value,p={site:$('line_site').value,legal_entity:$('line_legal_entity').value,account_number:$('line_account').value,landline:$('line_landline').value,provider:(ALL.find(x=>String(x.id)===String(id))?.provider||'WE'),service_type:$('line_service').value,package:$('line_package').value,status:$('line_status').value,speed_mbps:$('line_speed').value,quota_gb:$('line_quota').value,used_gb:$('line_used').value,renewal_cost_egp:$('line_renewal_cost').value,outstanding_egp:$('line_outstanding').value,due_date:$('line_due').value,renewal_date:$('line_renewal').value};const r=await fetch(id?`/api/lines/${id}`:'/api/lines',{method:id?'PUT':'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});const j=await r.json();if(!r.ok)return toast(j.error||'Unable to save line');closeModal('lineModal');toast(id?'Line updated':'Line added');loadData()}
async function deleteLine(id){const x=ALL.find(y=>y.id===id);if(!confirm(`Delete line ${x?.landline||id}?`))return;const r=await fetch(`/api/lines/${id}`,{method:'DELETE'});if(r.ok){toast('Line deleted');loadData()}else toast('Unable to delete line')}
function updateAccountProviderUI(){const p=$('acc_provider')?.value||'WE';if($('acc_login_label'))$('acc_login_label').textContent=p==='Orange'?'Orange Home Wireless Mobile Number *':'WE Service Number (Login) *';if($('acc_password_label'))$('acc_password_label').textContent=p==='Orange'?'Orange Account Password':'My WE Password';if($('acc_portal_type'))$('acc_portal_type').value=p==='Orange'?'Home Wireless':'Internet';if($('acc_username'))$('acc_username').placeholder=p==='Orange'?'012XXXXXXXX':'0237487976';}
function openAccount(a=null){$('accountForm').reset();$('account_id').value=a?.id||'';$('accountTitle').textContent=a?'Edit Internet Account':'Add Internet Account';$('acc_provider').value=a?.provider||'WE';$('acc_name').value=a?.name||'';$('acc_number').value=a?.account_number||'';$('acc_site').value=a?.site||'';$('acc_legal_entity').value=a?.legal_entity||'';$('acc_username').value=a?.login_username||'';$('acc_portal_type').value=a?.portal_service_type||'Internet';$('acc_enabled').value=String(a?.enabled??1);$('acc_password').value='';updateAccountProviderUI();openModal('accountModal')}
function editAccount(id){openAccount((DATA.accounts||[]).find(a=>a.id===id))}
async function saveAccount(e){e.preventDefault();const id=$('account_id').value,p={provider:$('acc_provider').value,name:$('acc_name').value,account_number:$('acc_username').value,site:$('acc_site').value,legal_entity:$('acc_legal_entity').value,login_username:$('acc_username').value,portal_service_type:$('acc_portal_type').value,password:$('acc_password').value,enabled:Number($('acc_enabled').value)};const r=await fetch(id?`/api/accounts/${id}`:'/api/accounts',{method:id?'PUT':'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});const j=await r.json();if(!r.ok)return toast(j.error||'Unable to save account');closeModal('accountModal');toast(id?'Account updated':'Internet account added');loadData()}
async function deleteAccount(id){const a=(DATA.accounts||[]).find(x=>x.id===id);if(!confirm(`Delete internet account ${a?.name||a?.login_username||id}? Lines will remain in the dashboard.`))return;const r=await fetch(`/api/accounts/${id}`,{method:'DELETE'});if(r.ok){toast('Account deleted');loadData()}else toast('Unable to delete account')}
async function syncLine(id){const b=$('lineSyncBtn-'+id);if(b){b.disabled=true;b.textContent='Syncing...'}try{const r=await fetch('/api/sync/line/'+id,{method:'POST'});const j=await r.json();if(!r.ok){toast(j.error||'Unable to start line sync');return}toast(j.message||'Line sync started');for(let i=0;i<120;i++){await new Promise(r=>setTimeout(r,2000));const sr=await fetch('/api/sync/status?_='+Date.now(),{cache:'no-store'});const st=await sr.json();if(st.running)continue;await loadData();if(st.last_error)toast(st.last_error);else toast(`Line sync complete: ${st.updated||0} line(s) updated`);return}toast('Line sync continues in the background')}catch(e){toast('Line sync request failed')}finally{if(b){b.disabled=false;b.textContent='Sync'}}}
async function stopBackground(){
  if(!confirm('Stop Internet Dashboard background service now?\n\nThe web page and hourly Auto Sync will stop until you start the background service again.'))return;
  const b=$('stopBgBtn');if(b){b.disabled=true;b.textContent='Stopping...'}
  try{const r=await fetch('/api/background/stop',{method:'POST'});const j=await r.json();if(!r.ok){toast(j.error||'Unable to stop background');if(b){b.disabled=false;b.textContent='Stop Background'};return}toast('Background service is stopping...');setTimeout(()=>{document.body.innerHTML='<div style=\"font-family:Segoe UI,Arial;padding:50px;text-align:center\"><h2>Internet Dashboard stopped</h2><p>The background service and hourly Auto Sync are now stopped.</p><p>Run <b>START_BACKGROUND_NOW.bat</b> to start it again.</p></div>'},700)}catch(e){toast('Background service is stopping...')}
}

async function syncNow(){const b=$('syncBtn');if(b){b.disabled=true;b.textContent='Starting Sync...'}try{const r=await fetch('/api/sync',{method:'POST'});const j=await r.json();if(!r.ok){toast(j.error||'Unable to start internet sync');return}toast(j.message||'Sync started');pollSync()}catch(e){toast('Internet sync request failed')}finally{if(b){b.disabled=false;b.textContent='Sync'}}}
async function pollSync(){for(let i=0;i<90;i++){await new Promise(r=>setTimeout(r,2000));try{const r=await fetch('/api/sync/status');const j=await r.json();const b=$('syncBtn');if(j.running){if(b)b.textContent='Syncing...';continue}if(b)b.textContent='Sync';await loadData();if(j.last_error)toast(j.last_error);else toast(`Sync complete: ${j.updated||0} line(s) updated`);return}catch(e){}}const b=$('syncBtn');if(b)b.textContent='Sync';toast('Internet sync is taking longer than expected; it continues in the background.')}
async function loadData(manual=false){const b=$('refreshBtn');if(manual&&b){b.disabled=true;b.textContent='Refreshing...'}try{const r=await fetch('/api/dashboard?_='+Date.now(),{cache:'no-store'});if(!r.ok)throw new Error('HTTP '+r.status);DATA=await r.json();ALL=(DATA.lines||[]).filter(x=>String(x.service_type||'').toLowerCase()!=='landline');const INTERNET_ACCOUNTS=(DATA.accounts||[]).filter(a=>String(a.portal_service_type||'Internet').toLowerCase()!=='landline');fillSelect('site',ALL.map(x=>x.site));fillSelect('legal_entity',ALL.map(x=>x.legal_entity));fillSelect('account_name',ALL.map(x=>x.account_name));$('accountList').innerHTML=INTERNET_ACCOUNTS.map(a=>`<option value="${esc(a.account_number)}">${esc(a.name||a.login_username||a.account_number||a.site||'')}</option>`).join('');$('sync').innerText='Last sync: '+DATA.last_sync;$('accounts').innerHTML=INTERNET_ACCOUNTS.map(a=>`<div class="account"><span><b>${esc(a.name||a.login_username||a.site||a.account_number||'Internet Account')}</b><br><small>${esc(a.provider||'WE')} · ${esc(a.login_username||'')} · ${a.lines||0} line(s) · ${esc(a.portal_service_type||'Internet')} · ${a.enabled?'Sync on':'Sync off'}${a.has_password?' · credentials saved':''}</small></span><span class="account-actions"><button class="btn small secondary" onclick="editAccount(${a.id})">Edit</button><button class="btn small danger" onclick="deleteAccount(${a.id})">Delete</button></span></div>`).join('')||'<div class="detail-empty">No internet accounts configured.</div>';render();renderNotifications();if(manual)toast('Dashboard refreshed')}catch(e){toast('Dashboard refresh failed: '+e.message)}finally{if(manual&&b){b.disabled=false;b.textContent='Refresh View'}}}
let SHARED_SYNC_RUNNING=false,SHARED_SYNC_FINISHED=null;
async function watchSharedSync(){try{const r=await fetch('/api/sync/status?_='+Date.now(),{cache:'no-store'});if(!r.ok)return;const j=await r.json();const b=$('syncBtn');if(j.running){SHARED_SYNC_RUNNING=true;if(b){b.disabled=true;b.textContent='Syncing...'}return}const finished=j.last_finished||null;const changed=(SHARED_SYNC_RUNNING&&finished)||(SHARED_SYNC_FINISHED&&finished&&finished!==SHARED_SYNC_FINISHED);if(b){b.disabled=false;b.textContent='Sync'}if(changed){await loadData();}SHARED_SYNC_RUNNING=false;if(finished)SHARED_SYNC_FINISHED=finished}catch(_){}}
initDetailPanel();loadData();watchSharedSync();setInterval(watchSharedSync,2000);setInterval(loadData,(parseInt('{{refresh}}')||60)*1000);
</script></body></html>'''


LANDLINES_HTML = r'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Internet Dashboard - Land Lines</title>
<style>
:root{--nav:#123d69;--nav2:#0b2d50;--blue:#174f83;--orange:#ff7900;--bg:#f3f7fb;--text:#17375e;--muted:#68778b;--line:#dde6ef;--ok:#18864b}
*{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:#18212f}.app{display:grid;grid-template-columns:195px 1fr;min-height:100vh}.side{background:linear-gradient(180deg,var(--nav),var(--nav2));color:#fff;position:sticky;top:0;height:100vh}.brand{padding:24px 18px 20px;font-size:22px;font-weight:800;line-height:1.05}.brand small{display:block;font-size:12px;margin-top:5px}.nav a{display:block;color:#eaf3ff;text-decoration:none;padding:15px 18px;border-left:4px solid transparent}.nav a:hover,.nav a.active{background:#ffffff14;border-left-color:var(--orange)}.ver{position:absolute;bottom:12px;left:14px;font-size:10px;opacity:.75}.main{min-width:0}.top{min-height:78px;background:linear-gradient(90deg,var(--blue),var(--nav));color:#fff;padding:13px 24px;display:flex;justify-content:space-between;align-items:center;gap:15px}.top h1{margin:0;font-size:27px}.sub{font-size:13px;color:#e4edf7;margin-top:4px}.actions{display:flex;gap:8px;flex-wrap:wrap;align-items:center}.btn{border:0;border-radius:9px;padding:10px 14px;font-weight:700;cursor:pointer;background:var(--blue);color:#fff}.btn.white{background:#fff;color:var(--text)}.btn.orange{background:var(--orange)}.wrap{padding:20px}.filters{display:grid;grid-template-columns:1.5fr 1fr 1fr 1fr auto;gap:10px;background:#fff;padding:12px;border:1px solid var(--line);border-radius:13px;margin-bottom:14px}.filters input,.filters select,.field input{width:100%;border:1px solid #dce5ef;border-radius:8px;padding:10px;background:#fff}.cards{display:grid;grid-template-columns:repeat(3,1fr);gap:12px;margin-bottom:14px}.card{background:#fff;border:1px solid var(--line);border-radius:13px;padding:16px}.k{font-size:12px;color:var(--muted)}.v{font-size:24px;font-weight:800;color:var(--text);margin-top:7px}.panel{background:#fff;border:1px solid var(--line);border-radius:13px;padding:16px}.land-layout{display:grid;grid-template-columns:minmax(0,2.6fr) minmax(320px,1fr);gap:14px}.land-layout.detail-hidden{grid-template-columns:1fr}.detail-panel.hidden{display:none}.detail-empty{color:var(--muted);font-size:13px;padding:35px 5px;text-align:center}.dhead{font-weight:800;font-size:18px;color:var(--text)}.dline{color:var(--muted);font-size:12px;margin:4px 0 15px}.dgrid{display:grid;grid-template-columns:1fr 1fr;gap:8px}.ditem{background:#f7f9fc;border-radius:9px;padding:10px}.ditem b{display:block;font-size:11px;color:var(--muted);font-weight:500}.ditem span{display:block;margin-top:5px;font-size:13px;font-weight:700}.panel-head{display:flex;justify-content:space-between;gap:10px;align-items:center;margin-bottom:10px}.panel h3{margin:0}.tablewrap{overflow:auto;max-height:650px}.tbl{width:100%;border-collapse:collapse;font-size:12.5px;white-space:nowrap}.tbl th,.tbl td{padding:10px 8px;border-bottom:1px solid #e8edf3;text-align:left}.tbl th{position:sticky;top:0;background:#fff;color:#667085}.pill{padding:4px 8px;border-radius:20px;font-size:11px;font-weight:700}.ok{background:#def7ea;color:#16784b}.warn{background:#fff0cf;color:#a76500}.bad{background:#fde2df;color:#b42318}.row-actions{display:flex;gap:5px}.btn.small{padding:7px 9px;font-size:11px}.modal{position:fixed;inset:0;background:#11182788;display:none;align-items:center;justify-content:center;padding:18px;z-index:100}.modal.open{display:flex}.modal-card{background:#fff;width:min(650px,100%);border-radius:14px;overflow:auto;max-height:92vh}.modal-head{display:flex;justify-content:space-between;align-items:center;padding:16px 18px;border-bottom:1px solid var(--line)}.x{border:0;background:transparent;font-size:24px;cursor:pointer}.form{padding:18px}.grid{display:grid;grid-template-columns:1fr 1fr;gap:12px}.field label{display:block;font-size:11px;color:var(--muted);margin-bottom:5px}.span2{grid-column:1/-1}.foot{display:flex;justify-content:flex-end;gap:8px;margin-top:18px}.toast{position:fixed;right:20px;bottom:20px;background:#111827;color:#fff;padding:11px 15px;border-radius:9px;display:none}.toast.show{display:block}@media(max-width:1000px){.app{grid-template-columns:1fr}.side{display:none}.cards{grid-template-columns:1fr}.filters{grid-template-columns:1fr 1fr}.top{align-items:flex-start;flex-direction:column}}
</style></head><body><div class="app"><aside class="side"><div class="brand">Internet Dashboard<small>WE & Orange</small></div><nav class="nav"><a href="/">Internet Lines</a><a class="active" href="/landlines">Land Lines</a><a href="/visual">Visual Dashboard</a><a href="/reports">Reports</a></nav><div class="ver">Unified Interface v5.33</div></aside><main class="main">
<header class="top"><div><h1>WE Land Lines</h1><div class="sub">WE fixed landline services — Portal Service Type: LandLine</div></div><div class="actions"><span id="syncState" class="sub"></span><button class="btn white" onclick="openAdd()">+ Add Land Line</button><button class="btn orange" id="syncAllBtn" onclick="syncAll()">Sync</button><button class="btn white" onclick="exportLandLines()">Export Excel</button><button class="btn white" onclick="loadData()">Refresh View</button></div></header>
<div class="wrap"><div class="filters"><input id="q" placeholder="Search line, branch, site, legal entity..." oninput="render()"><select id="legal" onchange="render()"><option value="">All Legal Entities</option></select><select id="branch" onchange="render()"><option value="">All Branches</option></select><select id="site" onchange="render()"><option value="">All Sites</option></select><button class="btn white" onclick="clearFilters()">Clear</button></div>
<div class="cards"><div class="card"><div class="k">Land Lines</div><div class="v" id="count">0</div></div><div class="card"><div class="k">Active</div><div class="v" id="active">0</div></div><div class="card"><div class="k">Pending / Attention</div><div class="v" id="pending">0</div></div></div>
<div class="land-layout" id="landLayout"><div class="panel"><div class="panel-head"><h3>WE Land Lines</h3><button class="btn small white" id="showDetailBtn" onclick="setDetailPanel(true)" style="display:none">Show Selected Line</button></div><div class="tablewrap"><table class="tbl"><thead><tr><th>Legal Entity</th><th>Site</th><th>Branch</th><th>Land Line</th><th>Package / Plan</th><th>Renewal Cost</th><th>Outstanding</th><th>Due</th><th>Renewal Date</th><th>Status</th><th>Last Sync</th><th>Manage</th></tr></thead><tbody id="rows"></tbody></table></div></div><div class="panel detail-panel" id="landDetailPanel"><div class="panel-head"><h3>Selected Line</h3><button class="btn small white" onclick="setDetailPanel(false)">Hide</button></div><div id="landDetail" class="detail-empty">Select any Land Line to show its full details here.</div></div></div></div>
</main></div>
<div class="modal" id="modal"><div class="modal-card"><div class="modal-head"><h2 id="modalTitle">Add WE Land Line</h2><button class="x" onclick="closeModal()">×</button></div><form class="form" onsubmit="save(event)"><input type="hidden" id="editAccountId"><div class="grid"><div class="field"><label>Legal Entity</label><input id="legalEntity" placeholder="Linah Farms"></div><div class="field"><label>Branch</label><input id="name" placeholder="Head Office"></div><div class="field"><label>Site</label><input id="siteInput" placeholder="HQ"></div><div class="field"><label>Service Type</label><input value="LandLine" disabled></div><div class="field span2"><label>WE Land Line / Service Number *</label><input id="number" placeholder="02XXXXXXXX" required></div><div class="field span2"><label id="passwordLabel">Password *</label><input id="password" type="password" required></div></div><div class="foot"><button type="button" class="btn white" onclick="closeModal()">Cancel</button><button class="btn orange" id="saveBtn" type="submit">Save Land Line</button></div></form></div></div><div class="toast" id="toast"></div>
<script>
let DATA={},ALL=[];const $=id=>document.getElementById(id);const esc=s=>String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));const money=v=>Number(v||0).toLocaleString(undefined,{maximumFractionDigits:2})+' EGP';
function toast(t){$('toast').textContent=t;$('toast').classList.add('show');setTimeout(()=>$('toast').classList.remove('show'),3500)}function fill(id,vals){const e=$(id),old=e.value,first=e.options[0].outerHTML;e.innerHTML=first+[...new Set(vals.filter(Boolean))].sort().map(v=>`<option>${esc(v)}</option>`).join('');e.value=[...e.options].some(o=>o.value===old)?old:''}
function filtered(){const q=$('q').value.trim().toLowerCase(),le=$('legal').value,br=$('branch').value,si=$('site').value;return ALL.filter(x=>{const h=[x.landline,x.legal_entity,x.account_name,x.site,x.package].join(' ').toLowerCase();return(!q||h.includes(q))&&(!le||x.legal_entity===le)&&(!br||x.account_name===br)&&(!si||x.site===si)})}
function pill(s){const c=s==='Active'?'ok':s==='High Usage'?'bad':'warn';return `<span class="pill ${c}">${esc(s||'-')}</span>`}
function render(){const a=filtered();$('count').textContent=a.length;$('active').textContent=a.filter(x=>x.status==='Active').length;$('pending').textContent=a.filter(x=>x.status!=='Active').length;$('rows').innerHTML=a.length?a.map(x=>`<tr onclick="showLandDetail(${x.id})"><td>${esc(x.legal_entity||'-')}</td><td>${esc(x.site||'-')}</td><td>${esc(x.account_name||'-')}</td><td><b>${esc(x.landline)}</b></td><td>${esc(x.package||'-')}</td><td>${money(x.renewal_cost_egp)}</td><td>${money(x.outstanding_egp)}</td><td>${esc(x.due_date||'-')}</td><td>${esc(x.renewal_date||'-')}</td><td>${pill(x.status)}</td><td>${esc(x.last_sync||'-')}</td><td><div class="row-actions"><button class="btn small" onclick="event.stopPropagation();syncLine(${x.id})">Sync</button><button class="btn small white" onclick="event.stopPropagation();editLine(${x.id})">Edit</button><button class="btn small" style="background:#b42318" onclick="event.stopPropagation();deleteLine(${x.id})">Delete</button></div></td></tr>`).join(''):'<tr><td colspan="12">No WE Land Lines found.</td></tr>'}
async function loadData(){const r=await fetch('/api/dashboard?_='+Date.now(),{cache:'no-store'});DATA=await r.json();ALL=(DATA.lines||[]).filter(x=>(x.provider||'WE')==='WE'&&String(x.service_type||'').toLowerCase()==='landline');fill('legal',ALL.map(x=>x.legal_entity));fill('branch',ALL.map(x=>x.account_name));fill('site',ALL.map(x=>x.site));const st=DATA.sync_state||{};$('syncState').textContent=st.running?'Syncing...':('Last sync: '+(DATA.last_sync||'Never'));$('syncAllBtn').disabled=!!st.running;render()}
function accountForLine(x){return (DATA.accounts||[]).find(a=>(a.provider||'WE')==='WE'&&String(a.portal_service_type||'').toLowerCase()==='landline'&&String(a.account_key||a.account_number||'')===String(x.account_number||''))}
function setDetailPanel(show){$('landDetailPanel').classList.toggle('hidden',!show);$('landLayout').classList.toggle('detail-hidden',!show);$('showDetailBtn').style.display=show?'none':'inline-block';localStorage.setItem('landDetailVisible',show?'1':'0')}
function showLandDetail(id){const x=ALL.find(y=>y.id===id);if(!x)return;$('landDetail').innerHTML=`<div class="dhead">${esc(x.site||'Land Line')}</div><div class="dline">${esc(x.landline)}</div><div class="dgrid"><div class="ditem"><b>Legal Entity</b><span>${esc(x.legal_entity||'-')}</span></div><div class="ditem"><b>Branch</b><span>${esc(x.account_name||'-')}</span></div><div class="ditem"><b>Provider</b><span>WE</span></div><div class="ditem"><b>Service</b><span>LandLine</span></div><div class="ditem"><b>Status</b><span>${pill(x.status)}</span></div><div class="ditem"><b>Package / Plan</b><span>${esc(x.package||'-')}</span></div><div class="ditem"><b>Renewal Cost</b><span>${money(x.renewal_cost_egp)}</span></div><div class="ditem"><b>Outstanding / Paid in Advance</b><span>${money(x.outstanding_egp)}</span></div><div class="ditem"><b>Due</b><span>${esc(x.due_date||'-')}</span></div><div class="ditem"><b>Renewal Date</b><span>${esc(x.renewal_date||'-')}</span></div><div class="ditem"><b>Source</b><span>${esc(x.source||'WE')}</span></div><div class="ditem"><b>Last Sync</b><span>${esc(x.last_sync||'-')}</span></div></div>`;setDetailPanel(true)}
function openAdd(){$('editAccountId').value='';$('modalTitle').textContent='Add WE Land Line';$('saveBtn').textContent='Save Land Line';$('passwordLabel').textContent='Password *';$('password').required=true;$('legalEntity').value='';$('name').value='';$('siteInput').value='';$('number').value='';$('password').value='';$('modal').classList.add('open')}
function editLine(id){const x=ALL.find(y=>y.id===id);if(!x)return;const a=accountForLine(x);if(!a){toast('Linked Land Line account was not found');return}$('editAccountId').value=a.id;$('modalTitle').textContent='Edit WE Land Line';$('saveBtn').textContent='Save Changes';$('passwordLabel').textContent='Password (leave blank to keep current)';$('password').required=false;$('legalEntity').value=a.legal_entity||x.legal_entity||'';$('name').value=a.name||x.account_name||'';$('siteInput').value=a.site||x.site||'';$('number').value=a.login_username||x.landline||'';$('password').value='';$('modal').classList.add('open')}
function closeModal(){$('modal').classList.remove('open')}function clearFilters(){['q','legal','branch','site'].forEach(id=>$(id).value='');render()}
function exportLandLines(){const p=new URLSearchParams();[['q','q'],['legal','legal_entity'],['branch','account_name'],['site','site']].forEach(([id,key])=>{const e=$(id);if(e&&e.value)p.set(key,e.value)});window.location.href='/export/landlines.xlsx'+(p.toString()?'?'+p.toString():'')}
async function save(e){e.preventDefault();const number=$('number').value.trim(),id=$('editAccountId').value;const p={provider:'WE',name:$('name').value.trim(),account_number:number,site:$('siteInput').value.trim(),legal_entity:$('legalEntity').value.trim(),login_username:number,portal_service_type:'LandLine',password:$('password').value,enabled:1};const r=await fetch(id?('/api/accounts/'+id):'/api/accounts',{method:id?'PUT':'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(p)});const d=await r.json();if(!r.ok){toast(d.error||(id?'Could not update Land Line':'Could not add Land Line'));return}closeModal();toast(id?'Land Line updated':'Land Line added. You can Sync it now.');e.target.reset();await loadData()}
async function syncAll(){const r=await fetch('/api/sync/landlines',{method:'POST'});const d=await r.json();if(!r.ok){toast(d.error||'Could not start Land Lines sync');return}toast('WE Land Lines sync started');poll()}async function syncLine(id){const r=await fetch('/api/sync/line/'+id,{method:'POST'});const d=await r.json();if(!r.ok){toast(d.error||'Could not start line sync');return}toast('Land Line sync started');poll()}function poll(){let n=0;const t=setInterval(async()=>{await loadData();n++;if(!(DATA.sync_state||{}).running||n>120)clearInterval(t)},2000)}
async function deleteLine(id){if(!confirm('Delete this Land Line from the dashboard?'))return;const r=await fetch('/api/lines/'+id,{method:'DELETE'});if(r.ok){toast('Land Line deleted');loadData()}else toast('Delete failed')}
setDetailPanel(localStorage.getItem('landDetailVisible')!=='0');loadData();setInterval(loadData,60000);
</script></body></html>'''

REPORTS_HTML = '<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Internet Dashboard — Reports</title>\n<style>\n:root{--we:#5b2d90;--we2:#7c4bb2;--bg:#f4f6fb;--card:#fff;--text:#18212f;--muted:#70798a;--line:#e8ebf2;--ok:#16784b;--warn:#a76500;--bad:#b42318;--blue:#1769aa}\n*{box-sizing:border-box}body{font-family:Segoe UI,Arial,sans-serif;background:var(--bg);margin:0;color:var(--text)}.wrap{max-width:1650px;margin:auto;padding:22px}.top{display:flex;justify-content:space-between;align-items:center;gap:14px;margin-bottom:16px}.title{font-size:27px;font-weight:750}.sub{color:var(--muted);font-size:13px;margin-top:4px}.actions{display:flex;gap:8px;align-items:center;flex-wrap:wrap}.btn{border:0;border-radius:9px;padding:10px 14px;background:var(--we);color:#fff;cursor:pointer;font-weight:600;text-decoration:none}.btn.outline{background:#fff;color:var(--we);border:1px solid #d9c9e9}.btn.secondary{background:#eef0f6;color:#333}.filters{display:grid;grid-template-columns:1.6fr 1fr 1fr 1fr 1fr auto;gap:10px;background:#fff;padding:12px;border-radius:13px;box-shadow:0 2px 10px #0000000a;margin-bottom:14px}.filters input,.filters select{border:1px solid #dce1ea;border-radius:8px;padding:10px;background:white;min-width:0;width:100%}.cards{display:grid;grid-template-columns:repeat(8,1fr);gap:10px;margin-bottom:14px}.card{background:#fff;border-radius:13px;padding:14px;box-shadow:0 2px 10px #0000000a}.k{font-size:11px;color:var(--muted)}.v{font-size:21px;font-weight:750;margin-top:6px}.mini{font-size:10px;color:var(--muted);margin-top:3px}.tabs{display:flex;gap:7px;flex-wrap:wrap;margin-bottom:12px}.tab{background:#fff;color:#555;border:1px solid #ddd;border-radius:9px;padding:9px 14px;cursor:pointer;font-weight:650}.tab.active{background:var(--we);color:#fff;border-color:var(--we)}.panel{background:#fff;border-radius:13px;padding:16px;box-shadow:0 2px 10px #0000000a}.panel-head{display:flex;justify-content:space-between;align-items:center;gap:10px;margin-bottom:10px}.panel h3{margin:0;font-size:16px}.tablewrap{overflow:auto;max-height:630px}.tbl{width:100%;border-collapse:collapse;font-size:12.5px;white-space:nowrap}.tbl th,.tbl td{padding:10px 8px;border-bottom:1px solid var(--line);text-align:left}.tbl th{position:sticky;top:0;background:#fff;color:#667085;z-index:2}.pill{padding:4px 8px;border-radius:20px;font-size:11px;font-weight:650;display:inline-block}.ok{background:#def7ea;color:var(--ok)}.warn{background:#fff0cf;color:var(--warn)}.bad{background:#fde2df;color:var(--bad)}.usagebar{width:100px;height:7px;background:#e8eaf0;border-radius:7px;overflow:hidden;display:inline-block;vertical-align:middle;margin-right:5px}.usagebar i{display:block;height:100%;background:var(--we)}.chartbox{display:grid;grid-template-columns:repeat(2,1fr);gap:12px;margin-bottom:14px}.chart{background:#fff;border-radius:13px;padding:16px;box-shadow:0 2px 10px #0000000a}.chart h3{font-size:14px;margin:0 0 12px}.barrow{display:grid;grid-template-columns:150px 1fr 85px;align-items:center;gap:8px;margin:8px 0;font-size:11px}.bar{height:9px;background:#eee;border-radius:9px;overflow:hidden}.bar i{display:block;height:100%;background:var(--we)}.empty{padding:35px;text-align:center;color:var(--muted)}\n@media(max-width:1250px){.cards{grid-template-columns:repeat(4,1fr)}.filters{grid-template-columns:1fr 1fr 1fr}.chartbox{grid-template-columns:1fr}}@media(max-width:700px){.wrap{padding:12px}.top{align-items:flex-start;flex-direction:column}.cards{grid-template-columns:repeat(2,1fr)}.filters{grid-template-columns:1fr}}\n/* Unified application shell matching Visual Dashboard */\n.app-shell{display:grid;grid-template-columns:195px 1fr;min-height:100vh}.side-shell{background:linear-gradient(180deg,#123d69,#0b2d50);color:#fff;padding:18px 0;position:sticky;top:0;height:100vh}.side-brand{padding:0 18px 20px;font-weight:800;font-size:22px;line-height:1.05}.side-brand small{display:block;font-size:12px;font-weight:600;opacity:.9;margin-top:5px}.side-nav a{display:block;padding:15px 18px;color:#eaf3ff;text-decoration:none;border-left:4px solid transparent;font-size:14px}.side-nav a.active,.side-nav a:hover{background:#ffffff14;border-left-color:#ff9c40}.side-ver{position:absolute;bottom:12px;left:14px;font-size:10px;opacity:.75}.main-shell{min-width:0}.main-shell>.wrap{max-width:none;margin:0;padding:0 20px 30px}.main-shell .top{min-height:78px;background:linear-gradient(90deg,#174f83,#123d69);margin:0 -20px 20px;padding:13px 24px;color:white}.main-shell .title{color:white}.main-shell .sub{color:#e4edf7}.main-shell .top .btn.outline,.main-shell .top .btn.secondary{background:#fff;color:#16345c;border-color:#d5e1ec}.panel,.card,.filters,.chart{border:1px solid #dce5ef;box-shadow:0 4px 18px #0f27430c}.tab.active{background:#174f83;border-color:#174f83}.bar i,.usagebar i{background:linear-gradient(90deg,#ff7900,#e9473f)}\n@media(max-width:1100px){.app-shell{grid-template-columns:1fr}.side-shell{display:none}.main-shell>.wrap{padding:0 12px 22px}.main-shell .top{margin:0 -12px 15px}}\n</style></head><body><div class="app-shell"><aside class="side-shell"><div class="side-brand">Internet Dashboard<small>WE & Orange</small></div><nav class="side-nav"><a href="/">Internet Lines</a><a href="/landlines">Land Lines</a><a href="/visual">Visual Dashboard</a><a class="active" href="/reports">Reports</a></nav><div class="side-ver">Unified Interface v5.33</div></aside><main class="main-shell"><div class="wrap">\n<div class="top"><div><div class="title">Internet Dashboard — Reports</div><div class="sub">Lines, consumption, cost and renewal reports</div></div><div class="actions"><button class="btn" onclick="exportReport()">Export Report Excel</button><button class="btn secondary" onclick="loadData()">Refresh</button></div></div>\n<div class="filters"><input id="q" placeholder="Search line, branch, site, package..." oninput="render()"><select id="legal_entity" onchange="render()"><option value="">All Legal Entities</option></select><select id="account_name" onchange="render()"><option value="">All Branches</option></select><select id="site" onchange="render()"><option value="">All Sites</option></select><select id="provider" onchange="render()"><option value="">All Providers</option><option value="WE">WE</option><option value="Orange">Orange Home Wireless</option></select><select id="status" onchange="render()"><option value="">All Status</option><option>Active</option><option>High Usage</option><option>Pending Sync</option></select><button class="btn secondary" onclick="clearFilters()">Clear</button></div>\n<div class="chartbox"><div class="chart"><h3>Usage by Branch</h3><div id="usageChart"></div></div><div class="chart"><h3>Renewal Cost by Legal Entity</h3><div id="costChart"></div></div></div>\n<div class="tabs"><button class="tab active" data-tab="lines" onclick="setTab(\'lines\',this)">Lines</button><button class="tab" data-tab="usage" onclick="setTab(\'usage\',this)">Usage</button><button class="tab" data-tab="cost" onclick="setTab(\'cost\',this)">Cost</button><button class="tab" data-tab="renewals" onclick="setTab(\'renewals\',this)">Renewals</button></div>\n<div class="panel"><div class="panel-head"><h3 id="reportTitle">Lines Report</h3><span id="countText" class="sub"></span></div><div class="tablewrap"><table class="tbl"><thead id="thead"></thead><tbody id="tbody"></tbody></table></div></div>\n</div></main></div><script>\nlet DATA={},ALL=[],TAB=\'lines\'; const $=id=>document.getElementById(id);const esc=s=>String(s??\'\').replace(/[&<>"\']/g,m=>({\'&\':\'&amp;\',\'<\':\'&lt;\',\'>\':\'&gt;\',\'"\':\'&quot;\',"\'":\'&#39;\'}[m]));const num=v=>Number(v||0);const fmt=v=>num(v).toLocaleString(undefined,{maximumFractionDigits:2});const money=v=>fmt(v)+\' EGP\';\nfunction fillSelect(id,vals){const el=$(id),old=el.value;const first=el.options[0].outerHTML;el.innerHTML=first+[...new Set(vals.filter(Boolean))].sort().map(v=>`<option>${esc(v)}</option>`).join(\'\');el.value=[...el.options].some(o=>o.value===old)?old:\'\'}\nfunction filtered(){const q=$(\'q\').value.trim().toLowerCase(),le=$(\'legal_entity\').value,br=$(\'account_name\').value,site=$(\'site\').value,pv=$(\'provider\').value,status=$(\'status\').value;return ALL.filter(x=>{const hay=[x.landline,x.legal_entity,x.account_name,x.site,x.package,x.service_type,x.status,x.provider].join(\' \').toLowerCase();return(!q||hay.includes(q))&&(!le||x.legal_entity===le)&&(!br||x.account_name===br)&&(!site||x.site===site)&&(!pv||x.provider===pv)&&(!status||x.status===status)})}\nfunction clearFilters(){[\'q\',\'legal_entity\',\'account_name\',\'site\',\'provider\',\'status\'].forEach(id=>$(id).value=\'\');render()}\nfunction setTab(t,b){TAB=t;document.querySelectorAll(\'.tab\').forEach(x=>x.classList.remove(\'active\'));b.classList.add(\'active\');renderTable(filtered())}\nfunction statusPill(s){const c=s===\'Active\'?\'ok\':s===\'High Usage\'?\'bad\':\'warn\';return `<span class="pill ${c}">${esc(s||\'-\')}</span>`}\nfunction renderCards(lines){}\nfunction group(lines,key,valueKey,avg=false){const m={};lines.forEach(x=>{const k=x[key]||\'Unassigned\';if(!m[k])m[k]=[];m[k].push(num(x[valueKey]))});return Object.entries(m).map(([name,arr])=>({name,value:avg?arr.reduce((a,b)=>a+b,0)/arr.length:arr.reduce((a,b)=>a+b,0)})).sort((a,b)=>b.value-a.value).slice(0,10)}\nfunction renderBars(id,rows,suffix){const max=Math.max(1,...rows.map(x=>x.value));$(id).innerHTML=rows.length?rows.map(x=>`<div class=barrow><span title="${esc(x.name)}">${esc(x.name)}</span><div class=bar><i style="width:${Math.min(100,x.value/max*100)}%"></i></div><b>${fmt(x.value)}${suffix}</b></div>`).join(\'\'):\'<div class=empty>No data</div>\'}\nfunction renderCharts(lines){renderBars(\'usageChart\',group(lines,\'account_name\',\'usage_pct\',true),\'%\');renderBars(\'costChart\',group(lines,\'legal_entity\',\'renewal_cost_egp\',false),\' EGP\')}\nconst COLS={lines:[[\'legal_entity\',\'Legal Entity\'],[\'account_name\',\'Branch\'],[\'site\',\'Site\'],[\'landline\',\'Line\'],[\'provider\',\'Provider\'],[\'package\',\'Package\'],[\'status\',\'Status\'],[\'last_sync\',\'Last Sync\']],usage:[[\'legal_entity\',\'Legal Entity\'],[\'account_name\',\'Branch\'],[\'landline\',\'Line\'],[\'provider\',\'Provider\'],[\'package\',\'Package\'],[\'quota_gb\',\'Quota GB\'],[\'used_gb\',\'Used GB\'],[\'remaining_gb\',\'Remaining GB\'],[\'usage_pct\',\'Usage %\'],[\'status\',\'Status\']],cost:[[\'legal_entity\',\'Legal Entity\'],[\'account_name\',\'Branch\'],[\'landline\',\'Line\'],[\'provider\',\'Provider\'],[\'renewal_cost_egp\',\'Renewal Cost\'],[\'outstanding_egp\',\'Outstanding\'],[\'charge_gap\',\'Charge Gap\'],[\'status\',\'Status\']],renewals:[[\'legal_entity\',\'Legal Entity\'],[\'account_name\',\'Branch\'],[\'landline\',\'Line\'],[\'provider\',\'Provider\'],[\'due_date\',\'Remaining Days\'],[\'renewal_date\',\'Renewal Date\'],[\'renewal_cost_egp\',\'Renewal Cost\'],[\'outstanding_egp\',\'Outstanding\'],[\'status\',\'Status\']]};\nfunction cell(x,k){if(k===\'status\')return statusPill(x[k]);if([\'quota_gb\',\'used_gb\',\'remaining_gb\'].includes(k))return fmt(x[k]);if(k===\'usage_pct\')return `<div class=usagebar><i style="width:${Math.min(100,num(x[k]))}%"></i></div>${num(x[k]).toFixed(1)}%`;if([\'renewal_cost_egp\',\'outstanding_egp\',\'charge_gap\'].includes(k))return money(x[k]);return esc(x[k]||\'-\')}\nfunction renderTable(lines){const titles={lines:\'Lines Report\',usage:\'Usage Report\',cost:\'Cost Report\',renewals:\'Renewals Report\'};$(\'reportTitle\').textContent=titles[TAB];$(\'countText\').textContent=lines.length+\' line(s)\';const cols=COLS[TAB];$(\'thead\').innerHTML=\'<tr>\'+cols.map(c=>`<th>${c[1]}</th>`).join(\'\')+\'</tr>\';const rows=lines.map(x=>({...x,charge_gap:Math.max(0,num(x.renewal_cost_egp)-num(x.outstanding_egp))}));if(TAB===\'usage\')rows.sort((a,b)=>num(b.usage_pct)-num(a.usage_pct));if(TAB===\'cost\')rows.sort((a,b)=>num(b.renewal_cost_egp)-num(a.renewal_cost_egp));$(\'tbody\').innerHTML=rows.length?rows.map(x=>\'<tr>\'+cols.map(c=>`<td>${cell(x,c[0])}</td>`).join(\'\')+\'</tr>\').join(\'\'):`<tr><td colspan="${cols.length}"><div class=empty>No matching data</div></td></tr>`}\nfunction render(){const lines=filtered();renderCharts(lines);renderTable(lines)}\nfunction queryParams(){const p=new URLSearchParams();[\'q\',\'legal_entity\',\'account_name\',\'site\',\'provider\',\'status\'].forEach(id=>{if($(id).value)p.set(id,$(id).value)});p.set(\'report\',TAB);return p}\nfunction exportReport(){window.location.href=\'/reports/export.xlsx?\'+queryParams().toString()}\nasync function loadData(){const r=await fetch(\'/api/dashboard?_=\'+Date.now(),{cache:\'no-store\'});DATA=await r.json();ALL=(DATA.lines||[]).filter(x=>String(x.service_type||\'\').toLowerCase()!==\'landline\');fillSelect(\'legal_entity\',ALL.map(x=>x.legal_entity));fillSelect(\'account_name\',ALL.map(x=>x.account_name));fillSelect(\'site\',ALL.map(x=>x.site));render()}\nloadData();setInterval(loadData,(parseInt(\'{{refresh}}\')||60)*1000);\n</script></body></html>'

VISUAL_HTML = r'''<!doctype html><html lang="en"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Internet Dashboard - Visual</title>
<style>
:root{--navy:#0f3a66;--navy2:#174f83;--bg:#f4f7fb;--card:#fff;--ink:#12345b;--muted:#6f8196;--line:#dbe5ef;--orange:#ff7900;--red:#ef4d43;--green:#0aa45a;--purple:#6b3aa8}
*{box-sizing:border-box}body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:#1d2d44}.app{display:grid;grid-template-columns:195px 1fr;min-height:100vh}.side{background:linear-gradient(180deg,#123d69,#0b2d50);color:white;position:sticky;top:0;height:100vh;padding:18px 0}.brand{padding:0 16px 20px;font-size:21px;font-weight:800;line-height:1}.brand small{display:block;font-size:12px;font-weight:600;opacity:.9;margin-top:6px}.nav a{display:block;padding:15px 16px;color:#edf5ff;text-decoration:none;border-left:4px solid transparent;font-size:14px}.nav a.active,.nav a:hover{background:#ffffff12;border-left-color:#ff9a32}.stop{position:absolute;bottom:44px;left:12px;right:12px;border:0;border-radius:8px;background:#7a2d59;color:white;padding:11px;font-weight:700;cursor:pointer}.ver{position:absolute;bottom:12px;left:14px;font-size:10px;opacity:.75}.main{min-width:0;padding:0 18px 30px}.top{margin:0 -18px 18px;padding:14px 22px;background:linear-gradient(90deg,#174f83,#123d69);color:white;display:flex;align-items:center;justify-content:space-between;gap:14px;min-height:76px}.top h1{margin:0;font-size:26px}.sub{font-size:12px;opacity:.85;margin-top:4px}.top-right{display:flex;align-items:center;gap:8px;flex-wrap:wrap}.pill{background:#ffffff15;padding:8px 10px;border-radius:9px;font-size:12px}.btn{border:0;border-radius:9px;padding:10px 13px;font-weight:700;cursor:pointer}.btn.white{background:#fff;color:#173d68}.btn.orange{background:var(--orange);color:white}.filters{display:grid;grid-template-columns:1.6fr repeat(5,1fr) auto;gap:9px;background:white;border:1px solid var(--line);padding:12px;border-radius:14px;margin-bottom:14px}.filters input,.filters select{border:1px solid #cfd9e5;border-radius:8px;padding:10px;background:white;min-width:0;width:100%}.section-head{display:flex;align-items:center;justify-content:space-between;margin:6px 2px 12px}.section-head h2{margin:0;color:var(--ink);font-size:18px}.section-head span{font-size:12px;color:var(--muted)}.cards-list{display:flex;flex-direction:column;gap:14px}.line-card{background:white;border:1px solid var(--line);border-radius:18px;box-shadow:0 5px 20px #0d29450a;padding:18px;display:grid;grid-template-columns:390px 1fr;gap:22px}.identity{padding-right:18px;border-right:1px solid #e7edf4}.linehead{display:flex;align-items:center;gap:18px;margin-bottom:18px}.logo{width:124px;height:124px;border-radius:15px;display:flex;align-items:center;justify-content:center;color:white;font-size:30px;font-weight:900;background:var(--orange);overflow:hidden}.logo.we{background:linear-gradient(135deg,#6b2b93,#8a45b8);font-size:38px}.line-title{font-size:28px;font-weight:800;color:#11355f}.line-num{font-size:25px;color:#2a5e91;margin-top:6px}.status{display:inline-block;margin-left:10px;padding:5px 11px;border-radius:18px;background:#dff7e9;color:#087b44;font-size:13px;font-weight:800;vertical-align:middle}.status.warn{background:#fff0cf;color:#a56500}.kv{display:grid;grid-template-columns:145px 1fr;gap:10px 12px;font-size:14px}.k{color:var(--muted)}.v{font-weight:700;color:#183654;overflow-wrap:anywhere}.metrics{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}.metric{border:1px solid var(--line);border-radius:14px;padding:20px 22px;background:#fff;min-height:125px}.metric .lab{font-size:13px;color:var(--muted)}.metric .val{font-size:30px;font-weight:900;margin-top:12px;color:#173e73}.metric.used .val{color:var(--red)}.metric.rem .val{color:var(--green)}.metric.outstanding .val{color:#173e73}.metric.outstanding.alert{border-color:#ef4d43;background:#fff4f3}.metric.outstanding.alert .lab,.metric.outstanding.alert .val{color:#d92d20}.usage{margin-top:16px;border:1px solid var(--line);border-radius:14px;padding:18px 20px}.usage-row{display:flex;justify-content:space-between;gap:10px;font-size:15px;font-weight:800;color:#1d3655}.bar{height:15px;border-radius:15px;background:#e5eaf1;overflow:hidden;margin-top:12px}.bar i{display:block;height:100%;background:linear-gradient(90deg,#ff7900,#ef4d43)}.empty{text-align:center;color:var(--muted);padding:45px;background:white;border:1px solid var(--line);border-radius:16px}.read-only{font-size:11px;color:var(--muted);margin:0 2px 12px}.provider-tag{display:inline-block;margin-top:3px;font-size:11px;color:#6a7c90}
@media(max-width:1350px){.line-card{grid-template-columns:330px 1fr}.metrics{grid-template-columns:repeat(2,1fr)}.logo{width:105px;height:105px}.line-title{font-size:23px}.line-num{font-size:21px}.filters{grid-template-columns:1fr 1fr 1fr 1fr}}
@media(max-width:1050px){.app{grid-template-columns:1fr}.side{display:none}.main{padding:0 12px 24px}.top{margin:0 -12px 15px}.line-card{grid-template-columns:1fr}.identity{border-right:0;border-bottom:1px solid #e7edf4;padding:0 0 16px}.filters{grid-template-columns:1fr 1fr 1fr}}
@media(max-width:650px){.metrics{grid-template-columns:1fr}.filters{grid-template-columns:1fr}.linehead{align-items:flex-start}.logo{width:82px;height:82px;font-size:22px}.line-title{font-size:20px}.line-num{font-size:18px}.kv{grid-template-columns:120px 1fr}}
</style></head><body><div class="app"><aside class="side"><div class="brand">Internet Dashboard<small>WE & Orange</small></div><nav class="nav"><a href="/">Internet Lines</a><a href="/landlines">Land Lines</a><a class="active" href="/visual">Visual Dashboard</a><a href="/reports">Reports</a></nav><div class="ver">Unified Interface v5.33</div></aside><main class="main"><header class="top"><div><h1>Internet Dashboard</h1><div class="sub">Visual Dashboard — All Lines</div></div><div class="top-right"><span class="pill" id="lastSync">Last sync: -</span><button class="btn white" id="visualRefreshBtn" onclick="refreshVisual()">Refresh View</button><button class="btn orange" onclick="exportExcel()">Export Excel</button></div></header>
<div class="filters"><input id="q" placeholder="Search line, branch, site, legal entity, package..." oninput="renderCards()"><select id="provider" onchange="renderCards()"><option value="">All Providers</option><option value="WE">WE</option><option value="Orange">Orange Home Wireless</option></select><select id="legal" onchange="renderCards()"><option value="">All Legal Entities</option></select><select id="branch" onchange="renderCards()"><option value="">All Branches</option></select><select id="site" onchange="renderCards()"><option value="">All Sites</option></select><select id="status" onchange="renderCards()"><option value="">All Status</option><option>Active</option><option>High Usage</option><option>Pending Sync</option></select><button class="btn white" onclick="clearFilters()">Clear</button></div>
<div class="read-only">Visual Dashboard reads the latest saved data from the main dashboard database.</div>
<div class="section-head"><h2>All Lines</h2><span id="visibleCount"></span></div><div id="cardsList" class="cards-list"></div>
</main></div><script>
let DATA=null,ALL=[];const $=id=>document.getElementById(id);const num=v=>Number(v||0);const fmt=(v,d=2)=>num(v).toLocaleString(undefined,{minimumFractionDigits:0,maximumFractionDigits:d});const esc=v=>String(v??'').replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]));
function fill(id,vals){const e=$(id),cur=e.value,first=e.options[0].outerHTML;e.innerHTML=first+[...new Set(vals.filter(Boolean))].sort().map(v=>`<option>${esc(v)}</option>`).join('');e.value=[...e.options].some(o=>o.value===cur)?cur:''}
function filtered(){const q=$('q').value.trim().toLowerCase(),p=$('provider').value,l=$('legal').value,b=$('branch').value,s=$('site').value,st=$('status').value;return ALL.filter(x=>{const hay=[x.landline,x.account_number,x.account_name,x.legal_entity,x.site,x.package,x.provider,x.status,x.service_type].join(' ').toLowerCase();return(!q||hay.includes(q))&&(!p||x.provider===p)&&(!l||x.legal_entity===l)&&(!b||x.account_name===b)&&(!s||x.site===s)&&(!st||x.status===st)})}
function clearFilters(){['q','provider','legal','branch','site','status'].forEach(id=>$(id).value='');renderCards()}
function card(x){const orange=x.provider==='Orange';const pct=Math.max(0,Math.min(100,num(x.usage_pct)||(num(x.quota_gb)?num(x.used_gb)/num(x.quota_gb)*100:0)));const statusClass=(x.status==='Active')?'status':'status warn';const providerName=orange?'Orange Home Wireless':'WE Internet';const needCharge=num(x.renewal_cost_egp)>num(x.outstanding_egp);return `<article class="line-card"><section class="identity"><div class="linehead"><div class="logo ${orange?'':'we'}">${orange?'orange':'WE'}</div><div><div class="line-title">${providerName}</div><div class="line-num">${esc(x.landline||x.account_number||'-')} <span class="${statusClass}">${esc(x.status||'Active')}</span></div><div class="provider-tag">${esc(x.service_type||'Internet')}</div></div></div><div class="kv"><div class="k">Package</div><div class="v">${esc(x.package||'-')}</div><div class="k">Renewal Cost</div><div class="v">${fmt(x.renewal_cost_egp)} EGP</div><div class="k">Due</div><div class="v">${esc(x.due_date||'-')}</div><div class="k">Last Sync</div><div class="v">${esc(x.last_sync||'-')}</div></div></section><section><div class="metrics"><div class="metric"><div class="lab">Total Quota</div><div class="val">${fmt(x.quota_gb)} GB</div></div><div class="metric used"><div class="lab">Used</div><div class="val">${fmt(x.used_gb)} GB</div></div><div class="metric rem"><div class="lab">Remaining</div><div class="val">${fmt(x.remaining_gb)} GB</div></div><div class="metric outstanding ${needCharge?'alert':''}"><div class="lab">Outstanding</div><div class="val">${fmt(x.outstanding_egp)} EGP</div></div></div><div class="usage"><div class="usage-row"><span>Usage (${fmt(x.used_gb)} GB / ${fmt(x.quota_gb)} GB)</span><span>${fmt(pct)}%</span></div><div class="bar"><i style="width:${pct}%"></i></div></div></section></article>`}
function renderCards(){const rows=filtered();$('visibleCount').textContent=`${rows.length} of ${ALL.length} lines`;$('cardsList').innerHTML=rows.length?rows.map(card).join(''):'<div class="empty">No matching lines</div>'}
function exportExcel(){window.location.href='/export/lines.xlsx'}
async function refreshVisual(){const b=$('visualRefreshBtn');if(b){b.disabled=true;b.textContent='Refreshing...'}try{await loadData()}finally{if(b){b.disabled=false;b.textContent='Refresh View'}}}
async function loadData(){const r=await fetch('/api/dashboard?_='+Date.now(),{cache:'no-store'});DATA=await r.json();ALL=(DATA.lines||[]).filter(x=>String(x.service_type||'').toLowerCase()!=='landline');$('lastSync').textContent='Last sync: '+(DATA.last_sync||'-');fill('legal',ALL.map(x=>x.legal_entity));fill('branch',ALL.map(x=>x.account_name));fill('site',ALL.map(x=>x.site));renderCards()}
loadData();setInterval(loadData,(parseInt('{{refresh}}')||60)*1000);
</script></body></html>'''
@app.route('/visual')
@app.route('/visual/')
@app.route('/visual-dashboard')
def visual_dashboard():
    return render_template_string(VISUAL_HTML, refresh=os.getenv('DASHBOARD_REFRESH_SECONDS','60'))

@app.route('/version')
def version():
    return jsonify({'build': APP_BUILD, 'visual_url': '/visual', 'ok': True})


def conn():
    c = sqlite3.connect(DB, timeout=10)
    c.row_factory = sqlite3.Row
    c.execute('PRAGMA foreign_keys=ON')
    return c


def get_fernet():
    env_key = os.getenv('WE_DASHBOARD_SECRET_KEY', '').strip()
    if env_key:
        key = env_key.encode()
    elif os.path.exists(KEY_FILE):
        with open(KEY_FILE, 'rb') as f:
            key = f.read().strip()
    else:
        key = Fernet.generate_key()
        with open(KEY_FILE, 'wb') as f:
            f.write(key)
    return Fernet(key)


def encrypt_secret(value):
    if not value:
        return None
    return get_fernet().encrypt(value.encode()).decode()


def decrypt_secret(value):
    if not value:
        return ''
    try:
        return get_fernet().decrypt(value.encode()).decode()
    except (InvalidToken, ValueError):
        return ''


def purge_legacy_email_settings_from_backups():
    """Remove legacy SMTP/email configuration from all retained SQLite backups."""
    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        for name in os.listdir(BACKUP_DIR):
            if not name.lower().endswith('.db'):
                continue
            path=os.path.join(BACKUP_DIR,name)
            try:
                c=sqlite3.connect(path)
                c.execute('DROP TABLE IF EXISTS email_settings')
                c.commit(); c.close()
            except Exception as e:
                print(f'Legacy email-settings purge skipped for {name}: {e}')
    except Exception as e:
        print('Legacy email-settings backup purge skipped:', e)


def _migrate_account_uniqueness(c):
    """Allow the same number in separate provider/service namespaces."""
    try:
        unique_single_number = False
        for idx in c.execute("PRAGMA index_list('we_accounts')").fetchall():
            idx_name = idx[1]
            is_unique = bool(idx[2])
            if not is_unique:
                continue
            cols = [r[2] for r in c.execute(f"PRAGMA index_info('{idx_name}')").fetchall()]
            if cols == ['account_number']:
                unique_single_number = True
                break
        if unique_single_number:
            c.execute('ALTER TABLE we_accounts RENAME TO we_accounts_legacy_unique')
            c.execute("""CREATE TABLE we_accounts(
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                provider TEXT NOT NULL DEFAULT 'WE',
                name TEXT,
                account_number TEXT NOT NULL,
                site TEXT,
                legal_entity TEXT,
                login_username TEXT,
                portal_service_type TEXT NOT NULL DEFAULT 'Internet',
                password_enc TEXT,
                enabled INTEGER NOT NULL DEFAULT 1,
                created_at TEXT,
                updated_at TEXT)""")
            legacy_cols = {r[1] for r in c.execute("PRAGMA table_info(we_accounts_legacy_unique)")}
            provider_expr = "COALESCE(provider,'WE')" if 'provider' in legacy_cols else "'WE'"
            portal_expr = "COALESCE(portal_service_type,'Internet')" if 'portal_service_type' in legacy_cols else "'Internet'"
            legal_expr = "COALESCE(legal_entity,'')" if 'legal_entity' in legacy_cols else "''"
            c.execute(f"""INSERT INTO we_accounts(
                id,provider,name,account_number,site,legal_entity,login_username,portal_service_type,password_enc,enabled,created_at,updated_at)
                SELECT id,{provider_expr},name,account_number,site,{legal_expr},login_username,{portal_expr},password_enc,enabled,created_at,updated_at
                FROM we_accounts_legacy_unique""")
            c.execute('DROP TABLE we_accounts_legacy_unique')
        c.execute("""CREATE UNIQUE INDEX IF NOT EXISTS ux_we_accounts_service
                     ON we_accounts(provider, account_number, portal_service_type)""")
    except Exception as e:
        print('Account uniqueness migration skipped:', e)


def portal_type_for_line(service_type, provider='WE'):
    p = str(provider or 'WE').strip().lower()
    st = str(service_type or '').strip().lower()
    if p == 'orange':
        return 'Home Wireless'
    if st == 'landline':
        return 'LandLine'
    return 'Internet'


def public_service_number(number):
    """Return the user-facing service number from an internal account key."""
    v = str(number or '').strip()
    prefix = 'WE::LANDLINE::'
    if v.upper().startswith(prefix):
        return v[len(prefix):]
    return v


def account_storage_number(number, provider='WE', portal_type='Internet'):
    """Namespace LandLine accounts so they never collide with Internet accounts."""
    raw = public_service_number(number)
    p = str(provider or 'WE').strip().lower()
    pt = str(portal_type or 'Internet').strip().lower()
    if p == 'we' and pt == 'landline':
        return f'WE::LANDLINE::{raw}'
    return raw


def migrate_landline_account_namespace(c):
    """Move legacy WE LandLine accounts/lines to an internal service namespace."""
    try:
        rows = c.execute("SELECT id, account_number, login_username FROM we_accounts WHERE lower(COALESCE(provider,'WE'))='we' AND lower(COALESCE(portal_service_type,'Internet'))='landline'").fetchall()
        for r in rows:
            old = str(r['account_number'] or '').strip()
            raw = str(r['login_username'] or '').strip() or public_service_number(old)
            new = account_storage_number(raw, 'WE', 'LandLine')
            if old == new:
                continue
            c.execute('UPDATE we_accounts SET account_number=? WHERE id=?', (new, r['id']))
            c.execute("UPDATE lines SET account_number=? WHERE account_number=? AND lower(COALESCE(provider,'WE'))='we' AND lower(COALESCE(service_type,''))='landline'", (new, old))
        c.commit()
    except Exception as e:
        print('LandLine namespace migration skipped:', e)


def init_db():
    c = conn(); cur = c.cursor()
    cur.execute('''CREATE TABLE IF NOT EXISTS we_accounts(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        provider TEXT NOT NULL DEFAULT 'WE',
        name TEXT,
        account_number TEXT NOT NULL,
        site TEXT,
        legal_entity TEXT,
        login_username TEXT,
        portal_service_type TEXT NOT NULL DEFAULT 'Internet',
        password_enc TEXT,
        enabled INTEGER NOT NULL DEFAULT 1,
        created_at TEXT,
        updated_at TEXT)''')
    account_cols = {r[1] for r in cur.execute('PRAGMA table_info(we_accounts)')}
    if 'provider' not in account_cols:
        cur.execute("ALTER TABLE we_accounts ADD COLUMN provider TEXT NOT NULL DEFAULT 'WE'")
    if 'portal_service_type' not in account_cols:
        cur.execute("ALTER TABLE we_accounts ADD COLUMN portal_service_type TEXT NOT NULL DEFAULT 'Internet'")
    if 'legal_entity' not in account_cols:
        cur.execute("ALTER TABLE we_accounts ADD COLUMN legal_entity TEXT DEFAULT ''")
    _migrate_account_uniqueness(c)
    migrate_landline_account_namespace(c)
    cur.execute('''CREATE TABLE IF NOT EXISTS sync_log(id INTEGER PRIMARY KEY AUTOINCREMENT, started_at TEXT, finished_at TEXT, updated INTEGER, failed INTEGER, message TEXT)''')
    cur.execute('''CREATE TABLE IF NOT EXISTS lines(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        site TEXT, legal_entity TEXT, account_number TEXT, landline TEXT, provider TEXT NOT NULL DEFAULT 'WE', service_type TEXT,
        package TEXT, speed_mbps INTEGER, quota_gb REAL, used_gb REAL,
        usage_pct REAL, remaining_gb REAL, bill_egp REAL, renewal_cost_egp REAL, outstanding_egp REAL,
        due_date TEXT, renewal_date TEXT, status TEXT, last_sync TEXT,
        source TEXT DEFAULT 'manual')''')
    cols = {r[1] for r in cur.execute('PRAGMA table_info(lines)')}
    required = {'provider':"TEXT NOT NULL DEFAULT 'WE'",'legal_entity':'TEXT DEFAULT \'\'','account_number':'TEXT','service_type':'TEXT','quota_gb':'REAL','used_gb':'REAL','remaining_gb':'REAL','renewal_date':'TEXT','renewal_cost_egp':'REAL DEFAULT 0','source':"TEXT DEFAULT 'manual'"}
    for name, typ in required.items():
        if name not in cols:
            cur.execute(f'ALTER TABLE lines ADD COLUMN {name} {typ}')
    now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    seed_samples = os.getenv('SEED_SAMPLE_DATA','0').strip().lower() in ('1','true','yes')
    if seed_samples and cur.execute('SELECT COUNT(*) FROM we_accounts').fetchone()[0] == 0:
        accs = [('WE','Head Office WE','ACC-10001','Head Office','','','Internet',None,1,now,now),('WE','October Factory WE','ACC-20007','6th October Factory','','','Internet',None,1,now,now),('WE','Farm 1 WE','ACC-31004','Farm 1','','','Internet',None,1,now,now),('WE','Branch 1 WE','ACC-44002','Branch 1','','','Internet',None,1,now,now)]
        cur.executemany('INSERT OR IGNORE INTO we_accounts(provider,name,account_number,site,legal_entity,login_username,portal_service_type,password_enc,enabled,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)', accs)
    if seed_samples and cur.execute('SELECT COUNT(*) FROM lines').fetchone()[0] == 0:
        rows = [
          ('Head Office','','ACC-10001','02-2450-1001','WE','Internet + Landline','WE Space 250',100,250,170,68,80,850,0,0,'2026-09-20','2026-09-20','Active'),
          ('Head Office','','ACC-10001','02-2450-1002','WE','Landline','Business Voice',0,0,0,0,0,210,0,0,'2026-09-20','2026-09-20','Active'),
          ('6th October Factory','','ACC-20007','02-3830-2100','WE','Internet + Landline','WE Space 400',200,400,376,94,24,1200,0,1200,'2026-09-18','2026-09-18','High Usage'),
          ('6th October Factory','','ACC-20007','02-3830-2101','WE','Internet Backup','WE Space 140',70,140,45,32,95,650,0,0,'2026-09-18','2026-09-18','Active'),
          ('Farm 1','','ACC-31004','045-320-4411','WE','Internet + Landline','WE Space 140',70,140,76,54,64,650,0,0,'2026-09-22','2026-09-22','Active'),
          ('Branch 1','','ACC-44002','02-2511-3300','WE','Internet + Landline','Landline + DSL',30,100,20,20,80,475,0,475,'2026-09-15','2026-09-15','Suspended')]
        cur.executemany('''INSERT INTO lines(site,legal_entity,account_number,landline,provider,service_type,package,speed_mbps,quota_gb,used_gb,usage_pct,remaining_gb,bill_egp,renewal_cost_egp,outstanding_egp,due_date,renewal_date,status,last_sync,source) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', [r+(now,'sample') for r in rows])
    # Email delivery was removed in v5.7. Purge legacy SMTP settings/credentials.
    cur.execute('DROP TABLE IF EXISTS email_settings')
    cur.execute('''CREATE TABLE IF NOT EXISTS notification_state(
        alert_key TEXT PRIMARY KEY, active INTEGER NOT NULL DEFAULT 1, last_sent TEXT, updated_at TEXT, read_at TEXT)''')
    notif_cols = {r[1] for r in cur.execute('PRAGMA table_info(notification_state)')}
    if 'read_at' not in notif_cols:
        cur.execute('ALTER TABLE notification_state ADD COLUMN read_at TEXT')
    c.commit(); c.close()
    purge_legacy_email_settings_from_backups()


def nfloat(v):
    try: return float(v or 0)
    except (TypeError, ValueError): return 0.0


def line_payload(data):
    # WE is the source of truth. Do not recompute Remaining when WE already
    # supplied it because top-ups/add-ons can make Remaining != Quota - Used.
    quota = nfloat(data.get('quota_gb'))
    used = nfloat(data.get('used_gb'))
    raw_remaining = data.get('remaining_gb')
    remaining = nfloat(raw_remaining) if raw_remaining not in (None, '') else max(0.0, quota - used) if quota else 0.0
    raw_usage = data.get('usage_pct')
    usage = nfloat(raw_usage) if raw_usage not in (None, '') else (used / quota * 100 if quota else 0.0)
    return {
        'site': str(data.get('site','')).strip(), 'legal_entity': str(data.get('legal_entity','')).strip(), 'account_number': str(data.get('account_number','')).strip(), 'landline': str(data.get('landline','')).strip(),
        'provider': str(data.get('provider','WE')).strip() or 'WE', 'service_type': str(data.get('service_type','Internet + Landline')).strip(), 'package': str(data.get('package','')).strip(),
        'speed_mbps': int(nfloat(data.get('speed_mbps'))),
        'quota_gb': round(quota, 2), 'used_gb': round(used, 2), 'usage_pct': round(usage, 2),
        'remaining_gb': round(remaining, 2),
        'bill_egp': round(nfloat(data.get('bill_egp')), 2), 'renewal_cost_egp': round(nfloat(data.get('renewal_cost_egp')), 2),
        'outstanding_egp': round(nfloat(data.get('outstanding_egp')), 2), 'due_date': str(data.get('due_date','')).strip(), 'renewal_date': str(data.get('renewal_date','')).strip(),
        'status': str(data.get('status','Active')).strip() or 'Active'}


def upsert_synced_lines(account_number, items):
    c = conn(); cur = c.cursor(); now = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    for raw in items:
        raw['account_number'] = raw.get('account_number') or account_number
        if not raw.get('provider'):
            ar = cur.execute('SELECT provider FROM we_accounts WHERE account_number=? ORDER BY id LIMIT 1',(raw['account_number'],)).fetchone()
            raw['provider'] = (ar['provider'] if ar and ar['provider'] else 'WE')
        x = line_payload(raw)
        existing = cur.execute('SELECT id, legal_entity, provider FROM lines WHERE account_number=? AND landline=? AND lower(provider)=lower(?) AND lower(service_type)=lower(?)', (x['account_number'], x['landline'], x['provider'], x['service_type'])).fetchone()
        if not x.get('legal_entity'):
            if existing and existing['legal_entity']:
                x['legal_entity'] = existing['legal_entity']
            else:
                ar = cur.execute('SELECT legal_entity FROM we_accounts WHERE account_number=? AND lower(provider)=lower(?) AND lower(portal_service_type)=lower(?) LIMIT 1', (x['account_number'], x['provider'], portal_type_for_line(x['service_type'], x['provider']))).fetchone()
                x['legal_entity'] = (ar['legal_entity'] if ar and ar['legal_entity'] else '')
        vals = (x['site'],x['legal_entity'],x['account_number'],x['landline'],x['provider'],x['service_type'],x['package'],x['speed_mbps'],x['quota_gb'],x['used_gb'],x['usage_pct'],x['remaining_gb'],x['bill_egp'],x.get('renewal_cost_egp',0),x['outstanding_egp'],x['due_date'],x['renewal_date'],x['status'],now,str(x.get('provider') or 'WE').lower())
        if existing:
            cur.execute('''UPDATE lines SET site=?,legal_entity=?,account_number=?,landline=?,provider=?,service_type=?,package=?,speed_mbps=?,quota_gb=?,used_gb=?,usage_pct=?,remaining_gb=?,bill_egp=?,renewal_cost_egp=?,outstanding_egp=?,due_date=?,renewal_date=?,status=?,last_sync=?,source=? WHERE id=?''', vals+(existing['id'],))
        else:
            cur.execute('''INSERT INTO lines(site,legal_entity,account_number,landline,provider,service_type,package,speed_mbps,quota_gb,used_gb,usage_pct,remaining_gb,bill_egp,renewal_cost_egp,outstanding_egp,due_date,renewal_date,status,last_sync,source) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''', vals)
    c.commit(); c.close()


@app.after_request
def add_no_cache_headers(response):
    if request.path.startswith('/api/') or request.path in ('/','/visual','/visual/','/reports','/landlines','/landlines/'):
        response.headers['Cache-Control']='no-store, no-cache, must-revalidate, max-age=0'
        response.headers['Pragma']='no-cache'
        response.headers['Expires']='0'
    return response


@app.route('/')
def index():
    # Opening the web dashboard must not trigger an internet-provider sync.
    # Synchronization is handled only by the hourly background scheduler
    # or by the manual Sync Internet Now / per-line Sync actions.
    return render_template_string(HTML, refresh=os.getenv('DASHBOARD_REFRESH_SECONDS','60'))




@app.route('/landlines')
@app.route('/landlines/')
def landlines_page():
    return render_template_string(LANDLINES_HTML)

@app.route('/reports')
def reports_page():
    return render_template_string(REPORTS_HTML, refresh=os.getenv('DASHBOARD_REFRESH_SECONDS','60'))



EXPORT_COLUMNS = [
    ('legal_entity','Legal Entity'),
    ('account_name','Branch'),
    ('site','Site'),
    ('landline','Line Number'),
    ('provider','Provider'),
    ('service_type','Service Type'),
    ('package','Package'),
    ('speed_mbps','Speed (Mbps)'),
    ('quota_gb','Quota (GB)'),
    ('used_gb','Used (GB)'),
    ('usage_pct','Usage (%)'),
    ('remaining_gb','Remaining (GB)'),
    ('renewal_cost_egp','Renewal Cost (EGP)'),
    ('outstanding_egp','Outstanding / Current Balance (EGP)'),
    ('due_date','Due / Remaining Days'),
    ('renewal_date','Renewal Date'),
    ('status','Status'),
    ('last_sync','Last Sync'),
    ('source','Source'),
]

def _is_landline_row(x):
    return str(x.get('service_type') or '').strip().lower() == 'landline'

def _get_export_rows(scope='internet'):
    c=conn()
    rows=[dict(r) for r in c.execute('SELECT * FROM lines ORDER BY legal_entity,site,landline')]
    accounts=[dict(r) for r in c.execute('SELECT provider,name,account_number,portal_service_type FROM we_accounts')]
    c.close()
    names={(str(a.get('provider') or 'WE').lower(), a.get('account_number') or '', str(a.get('portal_service_type') or 'Internet').lower()): a.get('name') or '' for a in accounts}
    for x in rows:
        p=str(x.get('provider') or 'WE')
        pt=portal_type_for_line(x.get('service_type'), p)
        x['account_name']=names.get((p.lower(), x.get('account_number') or '', pt.lower()), '')
    if scope == 'landline':
        rows=[x for x in rows if _is_landline_row(x)]
    elif scope != 'all':
        rows=[x for x in rows if not _is_landline_row(x)]
    q=(request.args.get('q') or '').strip().lower()
    site=(request.args.get('site') or '').strip()
    legal=(request.args.get('legal_entity') or '').strip()
    branch=(request.args.get('account_name') or '').strip()
    provider=(request.args.get('provider') or '').strip()
    status=(request.args.get('status') or '').strip()
    out=[]
    for x in rows:
        hay=' '.join(str(x.get(k) or '') for k in ('site','legal_entity','account_name','landline','package','service_type')).lower()
        if q and q not in hay: continue
        if site and (x.get('site') or '') != site: continue
        if legal and (x.get('legal_entity') or '') != legal: continue
        if branch and (x.get('account_name') or '') != branch: continue
        if provider and (x.get('provider') or 'WE') != provider: continue
        if status and (x.get('status') or '') != status: continue
        out.append(x)
    return out

def _excel_value(v):
    return '' if v is None else v

@app.route('/export/lines.xlsx')
def export_lines_xlsx():
    rows=_get_export_rows()
    wb=Workbook(); ws=wb.active; ws.title='Internet Lines'
    title='Internet Dashboard - Lines Export'
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(EXPORT_COLUMNS))
    cell=ws.cell(1,1,title); cell.font=Font(bold=True,size=16,color='FFFFFF'); cell.fill=PatternFill('solid',fgColor='5B2D90'); cell.alignment=Alignment(horizontal='left')
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=len(EXPORT_COLUMNS))
    ws.cell(2,1,f'Exported: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} | Lines: {len(rows)}').font=Font(italic=True,color='666666')
    header_row=4
    for col,(_,label) in enumerate(EXPORT_COLUMNS,1):
        c=ws.cell(header_row,col,label); c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='7C4BB2'); c.alignment=Alignment(horizontal='center',vertical='center')
    for r_idx,row in enumerate(rows,header_row+1):
        for c_idx,(key,_) in enumerate(EXPORT_COLUMNS,1):
            ws.cell(r_idx,c_idx,_excel_value(row.get(key)))
    # formats
    money_keys={'renewal_cost_egp','outstanding_egp'}
    decimal_keys={'quota_gb','used_gb','remaining_gb'}
    pct_key='usage_pct'
    for c_idx,(key,_) in enumerate(EXPORT_COLUMNS,1):
        if key in money_keys:
            for r in range(header_row+1,header_row+1+len(rows)): ws.cell(r,c_idx).number_format='#,##0.00 "EGP"'
        elif key in decimal_keys:
            for r in range(header_row+1,header_row+1+len(rows)): ws.cell(r,c_idx).number_format='#,##0.00'
        elif key==pct_key:
            for r in range(header_row+1,header_row+1+len(rows)): ws.cell(r,c_idx).number_format='0.00"%"'
    ws.freeze_panes='A5'; ws.auto_filter.ref=f'A4:{get_column_letter(len(EXPORT_COLUMNS))}{header_row+len(rows)}'
    thin=Side(style='thin',color='E6E8EC')
    for row in ws.iter_rows(min_row=header_row,max_row=header_row+len(rows),min_col=1,max_col=len(EXPORT_COLUMNS)):
        for c in row: c.border=Border(bottom=thin); c.alignment=Alignment(vertical='center')
    for idx,(key,label) in enumerate(EXPORT_COLUMNS,1):
        width=max(len(label)+2,12)
        vals=[str(r.get(key) or '') for r in rows[:300]]
        if vals: width=min(max(width,max(map(len,vals))+2),36)
        ws.column_dimensions[get_column_letter(idx)].width=width
    ws.row_dimensions[1].height=26; ws.sheet_view.showGridLines=False
    bio=io.BytesIO(); wb.save(bio); bio.seek(0)
    fn=f'Internet_Lines_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    return send_file(bio,as_attachment=True,download_name=fn,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


LANDLINE_EXPORT_COLUMNS = [
    ('legal_entity','Legal Entity'),('site','Site'),('account_name','Branch'),('landline','Land Line'),
    ('provider','Provider'),('service_type','Service Type'),('package','Package / Plan'),
    ('renewal_cost_egp','Renewal Cost (EGP)'),('outstanding_egp','Outstanding / Paid in Advance (EGP)'),
    ('due_date','Due / Remaining Days'),('renewal_date','Renewal Date'),('status','Status'),('last_sync','Last Sync'),('source','Source')
]

@app.route('/export/landlines.xlsx')
def export_landlines_xlsx():
    rows=_get_export_rows('landline')
    wb=Workbook(); ws=wb.active; ws.title='WE Land Lines'
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=len(LANDLINE_EXPORT_COLUMNS))
    c=ws.cell(1,1,'Internet Dashboard - WE Land Lines Export'); c.font=Font(bold=True,size=16,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='174F83'); c.alignment=Alignment(horizontal='left')
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=len(LANDLINE_EXPORT_COLUMNS))
    ws.cell(2,1,f'Exported: {datetime.now().strftime("%Y-%m-%d %H:%M:%S")} | Land Lines: {len(rows)}').font=Font(italic=True,color='666666')
    header_row=4
    for col,(_,label) in enumerate(LANDLINE_EXPORT_COLUMNS,1):
        c=ws.cell(header_row,col,label); c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='FF7900'); c.alignment=Alignment(horizontal='center',vertical='center')
    for r_idx,row in enumerate(rows,header_row+1):
        for c_idx,(key,_) in enumerate(LANDLINE_EXPORT_COLUMNS,1): ws.cell(r_idx,c_idx,_excel_value(row.get(key)))
    for c_idx,(key,_) in enumerate(LANDLINE_EXPORT_COLUMNS,1):
        if key in {'renewal_cost_egp','outstanding_egp'}:
            for r in range(header_row+1,header_row+1+len(rows)): ws.cell(r,c_idx).number_format='#,##0.00 "EGP"'
    if rows: ws.auto_filter.ref=f'A4:{get_column_letter(len(LANDLINE_EXPORT_COLUMNS))}{header_row+len(rows)}'
    ws.freeze_panes='A5'; ws.sheet_view.showGridLines=False
    thin=Side(style='thin',color='E6E8EC')
    for row in ws.iter_rows(min_row=header_row,max_row=header_row+len(rows),min_col=1,max_col=len(LANDLINE_EXPORT_COLUMNS)):
        for c in row: c.border=Border(bottom=thin); c.alignment=Alignment(vertical='center')
    for idx,(key,label) in enumerate(LANDLINE_EXPORT_COLUMNS,1):
        vals=[str(r.get(key) or '') for r in rows[:300]]; width=max(len(label)+2,12)
        if vals: width=min(max(width,max(map(len,vals))+2),36)
        ws.column_dimensions[get_column_letter(idx)].width=width
    bio=io.BytesIO(); wb.save(bio); bio.seek(0)
    fn=f'WE_Land_Lines_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    return send_file(bio,as_attachment=True,download_name=fn,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


def _report_filtered_rows():
    return _get_export_rows('internet')

def _write_sheet(ws, title, columns, rows):
    ws.title=title[:31]
    headers=[label for _,label in columns]
    for c_idx,label in enumerate(headers,1):
        c=ws.cell(1,c_idx,label); c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='5B2D90'); c.alignment=Alignment(horizontal='center')
    for r_idx,row in enumerate(rows,2):
        enriched=dict(row)
        enriched['charge_gap']=max(0,nfloat(row.get('renewal_cost_egp'))-nfloat(row.get('outstanding_egp')))
        for c_idx,(key,_) in enumerate(columns,1):
            ws.cell(r_idx,c_idx,_excel_value(enriched.get(key)))
    ws.freeze_panes='A2'; ws.sheet_view.showGridLines=False
    if rows: ws.auto_filter.ref=f'A1:{get_column_letter(len(columns))}{len(rows)+1}'
    money={'renewal_cost_egp','outstanding_egp','charge_gap'}; decimals={'quota_gb','used_gb','remaining_gb'}
    for c_idx,(key,label) in enumerate(columns,1):
        if key in money:
            for r in range(2,len(rows)+2): ws.cell(r,c_idx).number_format='#,##0.00 "EGP"'
        elif key in decimals:
            for r in range(2,len(rows)+2): ws.cell(r,c_idx).number_format='#,##0.00'
        elif key=='usage_pct':
            for r in range(2,len(rows)+2): ws.cell(r,c_idx).number_format='0.00"%"'
        vals=[str((row.get(key) if key!='charge_gap' else max(0,nfloat(row.get('renewal_cost_egp'))-nfloat(row.get('outstanding_egp')))) or '') for row in rows[:300]]
        width=max(12,len(label)+2,max([len(v)+2 for v in vals] or [12])); ws.column_dimensions[get_column_letter(c_idx)].width=min(width,34)

@app.route('/reports/export.xlsx')
def export_reports_xlsx():
    rows=_report_filtered_rows(); report=(request.args.get('report') or 'all').strip().lower()
    wb=Workbook(); wb.remove(wb.active)
    sets={
      'lines': [('legal_entity','Legal Entity'),('account_name','Branch'),('site','Site'),('landline','Line Number'),('provider','Provider'),('package','Package'),('status','Status'),('last_sync','Last Sync')],
      'usage': [('legal_entity','Legal Entity'),('account_name','Branch'),('landline','Line Number'),('provider','Provider'),('package','Package'),('quota_gb','Quota (GB)'),('used_gb','Used (GB)'),('remaining_gb','Remaining (GB)'),('usage_pct','Usage (%)'),('status','Status')],
      'cost': [('legal_entity','Legal Entity'),('account_name','Branch'),('landline','Line Number'),('provider','Provider'),('renewal_cost_egp','Renewal Cost (EGP)'),('outstanding_egp','Outstanding (EGP)'),('charge_gap','Required Charge (EGP)'),('status','Status')],
      'renewals': [('legal_entity','Legal Entity'),('account_name','Branch'),('landline','Line Number'),('provider','Provider'),('due_date','Remaining Days'),('renewal_date','Renewal Date'),('renewal_cost_egp','Renewal Cost (EGP)'),('outstanding_egp','Outstanding (EGP)'),('status','Status')]
    }
    names={'lines':'Lines','usage':'Usage','cost':'Cost','renewals':'Renewals'}
    targets=[report] if report in sets else ['lines','usage','cost','renewals']
    for key in targets:
        data=list(rows)
        if key=='usage': data.sort(key=lambda x:nfloat(x.get('usage_pct')),reverse=True)
        elif key=='cost': data.sort(key=lambda x:nfloat(x.get('renewal_cost_egp')),reverse=True)
        _write_sheet(wb.create_sheet(),names[key],sets[key],data)
    # Summary sheet first
    ws=wb.create_sheet('Summary',0)
    total_quota=sum(nfloat(x.get('quota_gb')) for x in rows); total_used=sum(nfloat(x.get('used_gb')) for x in rows); total_rem=sum(nfloat(x.get('remaining_gb')) for x in rows); total_cost=sum(nfloat(x.get('renewal_cost_egp')) for x in rows); total_out=sum(nfloat(x.get('outstanding_egp')) for x in rows)
    summary=[('Metric','Value'),('Lines',len(rows)),('Total Quota (GB)',total_quota),('Used (GB)',total_used),('Remaining (GB)',total_rem),('Average Usage (%)',(sum(nfloat(x.get('usage_pct')) for x in rows)/len(rows) if rows else 0)),('Renewal Cost (EGP)',total_cost),('Outstanding (EGP)',total_out),('High Usage',sum(1 for x in rows if str(x.get('status') or '')=='High Usage' or nfloat(x.get('usage_pct'))>=90))]
    for r,row in enumerate(summary,1):
        for c,v in enumerate(row,1): ws.cell(r,c,v)
    for c in ws[1]: c.font=Font(bold=True,color='FFFFFF'); c.fill=PatternFill('solid',fgColor='5B2D90')
    ws.column_dimensions['A'].width=25; ws.column_dimensions['B'].width=20; ws.sheet_view.showGridLines=False
    bio=io.BytesIO(); wb.save(bio); bio.seek(0)
    fn=f'Internet_Reports_{datetime.now().strftime("%Y%m%d_%H%M%S")}.xlsx'
    return send_file(bio,as_attachment=True,download_name=fn,mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


@app.route('/api/dashboard')
def dashboard():
    c = conn(); rows = [dict(r) for r in c.execute('SELECT * FROM lines ORDER BY site,account_number,landline')]
    db_accounts = [dict(r) for r in c.execute('SELECT id,provider,name,account_number,site,legal_entity,login_username,portal_service_type,password_enc,enabled FROM we_accounts ORDER BY name,account_number')]
    account_names = {(str(a['provider'] or 'WE').lower(), a['account_number'], str(a['portal_service_type'] or 'Internet').lower()): (a['name'] or '') for a in db_accounts}
    account_providers = {(a['account_number'], str(a['portal_service_type'] or 'Internet').lower()): (a['provider'] or 'WE') for a in db_accounts}
    for x in rows:
        p = str(x.get('provider') or 'WE')
        pt = portal_type_for_line(x.get('service_type'), p)
        x['account_name'] = account_names.get((p.lower(), x.get('account_number') or '', pt.lower()), '')
        x['provider'] = x.get('provider') or account_providers.get((x.get('account_number') or '', pt.lower()), 'WE')
    accounts = []
    for a in db_accounts:
        rr = [x for x in rows if (x.get('account_number') or '') == a['account_number'] and str(x.get('provider') or 'WE').lower() == str(a['provider'] or 'WE').lower() and portal_type_for_line(x.get('service_type'), x.get('provider')).lower() == str(a['portal_service_type'] or 'Internet').lower()]
        display_name=(a['name'] or '').strip(); display_name = (a['login_username'] or public_service_number(a['account_number'])) if (not display_name or display_name.lower()=='internet account') else display_name
        accounts.append({'id':a['id'],'provider':a['provider'] or 'WE','name':display_name,'account_number':public_service_number(a['account_number']),'account_key':a['account_number'],'site':a['site'],'legal_entity':a['legal_entity'],'login_username':a['login_username'] or public_service_number(a['account_number']),'portal_service_type':a['portal_service_type'] or 'Internet','enabled':bool(a['enabled']),'has_password':bool(a['password_enc']),'lines':len(rr),'sites':len({x['site'] for x in rr if x['site']}),'monthly_bill':sum(nfloat(x['renewal_cost_egp']) for x in rr)})
    c.close(); last = max([x['last_sync'] for x in rows if x.get('last_sync')], default='Never')
    notifications = sync_notification_state(rows)
    return jsonify({'lines':rows,'accounts':accounts,'last_sync':last,'sync_state':SYNC_STATE,'notifications':notifications,'unread_notifications':sum(1 for a in notifications if not a.get('read'))})


@app.route('/api/lines', methods=['POST'])
def add_line():
    data = request.get_json(silent=True) or {}; x = line_payload(data)
    if not x['site'] or not x['account_number'] or not x['landline']: return jsonify({'error':'Site, account number and landline are required.'}),400
    c=conn(); now=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    if c.execute('SELECT 1 FROM lines WHERE account_number=? AND landline=? AND lower(provider)=lower(?) AND lower(service_type)=lower(?)',(x['account_number'],x['landline'],x['provider'],x['service_type'])).fetchone(): c.close(); return jsonify({'error':'This landline already exists under the selected account.'}),409
    cur=c.execute('''INSERT INTO lines(site,legal_entity,account_number,landline,provider,service_type,package,speed_mbps,quota_gb,used_gb,usage_pct,remaining_gb,bill_egp,renewal_cost_egp,outstanding_egp,due_date,renewal_date,status,last_sync,source) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)''',(x['site'],x['legal_entity'],x['account_number'],x['landline'],x['provider'],x['service_type'],x['package'],x['speed_mbps'],x['quota_gb'],x['used_gb'],x['usage_pct'],x['remaining_gb'],x['bill_egp'],x.get('renewal_cost_egp',0),x['outstanding_egp'],x['due_date'],x['renewal_date'],x['status'],now,'manual'))
    c.commit(); new_id=cur.lastrowid; c.close(); return jsonify({'ok':True,'id':new_id}),201


@app.route('/api/lines/<int:line_id>', methods=['PUT','DELETE'])
def manage_line(line_id):
    c=conn()
    if request.method=='DELETE':
        c.execute('DELETE FROM lines WHERE id=?',(line_id,)); c.commit(); c.close(); return jsonify({'ok':True})
    data=request.get_json(silent=True) or {}; x=line_payload(data)
    if not x['site'] or not x['account_number'] or not x['landline']: c.close(); return jsonify({'error':'Site, account number and landline are required.'}),400
    duplicate=c.execute('SELECT id FROM lines WHERE account_number=? AND landline=? AND lower(provider)=lower(?) AND lower(service_type)=lower(?) AND id<>?',(x['account_number'],x['landline'],x['provider'],x['service_type'],line_id)).fetchone()
    if duplicate: c.close(); return jsonify({'error':'This landline already exists under the selected account.'}),409
    now=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute('''UPDATE lines SET site=?,legal_entity=?,account_number=?,landline=?,provider=?,service_type=?,package=?,speed_mbps=?,quota_gb=?,used_gb=?,usage_pct=?,remaining_gb=?,bill_egp=?,renewal_cost_egp=?,outstanding_egp=?,due_date=?,renewal_date=?,status=?,last_sync=?,source='manual' WHERE id=?''',(x['site'],x['legal_entity'],x['account_number'],x['landline'],x['provider'],x['service_type'],x['package'],x['speed_mbps'],x['quota_gb'],x['used_gb'],x['usage_pct'],x['remaining_gb'],x['bill_egp'],x.get('renewal_cost_egp',0),x['outstanding_egp'],x['due_date'],x['renewal_date'],x['status'],now,line_id))
    c.commit(); c.close(); return jsonify({'ok':True})


@app.route('/api/accounts', methods=['POST'])
def add_account():
    data=request.get_json(silent=True) or {}; raw_number=str(data.get('account_number') or data.get('login_username') or '').strip()
    if not raw_number: return jsonify({'error':'Service/login number is required.'}),400
    now=datetime.now().strftime('%Y-%m-%d %H:%M:%S'); c=conn()
    try:
        provider=str(data.get('provider','WE')).strip() or 'WE'; name=str(data.get('name','')).strip(); site=str(data.get('site','')).strip(); legal_entity=str(data.get('legal_entity','')).strip(); login_username=str(data.get('login_username','')).strip() or raw_number; portal_type=str(data.get('portal_service_type','Internet')).strip() or 'Internet'
        number=account_storage_number(raw_number, provider, portal_type)
        dup=c.execute('SELECT id FROM we_accounts WHERE lower(provider)=lower(?) AND account_number=? AND lower(portal_service_type)=lower(?)',(provider,number,portal_type)).fetchone()
        if dup:
            c.close(); return jsonify({'error':f'This {portal_type} service number already exists.'}),409
        cur=c.execute('''INSERT INTO we_accounts(provider,name,account_number,site,legal_entity,login_username,portal_service_type,password_enc,enabled,created_at,updated_at) VALUES(?,?,?,?,?,?,?,?,?,?,?)''',(provider,name,number,site,legal_entity,login_username,portal_type,encrypt_secret(str(data.get('password',''))) if data.get('password') else None,1 if data.get('enabled',1) else 0,now,now)); c.commit(); new_id=cur.lastrowid
        if login_username:
            exists=c.execute('SELECT id FROM lines WHERE account_number=? AND landline=? AND lower(provider)=lower(?) AND lower(service_type)=lower(?)',(number,login_username,provider,'Home Wireless' if provider.lower()=='orange' else ('LandLine' if portal_type.lower()=='landline' else 'Internet'))).fetchone()
            if not exists:
                sql='INSERT INTO lines(site,legal_entity,account_number,landline,provider,service_type,package,speed_mbps,quota_gb,used_gb,usage_pct,remaining_gb,bill_egp,renewal_cost_egp,outstanding_egp,due_date,renewal_date,status,last_sync,source) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
                c.execute(sql,(site,legal_entity,number,login_username,provider,'Home Wireless' if provider.lower()=='orange' else ('LandLine' if portal_type.lower()=='landline' else 'Internet'),'',0,0,0,0,0,0,0,0,'','','Pending Sync',now,f'{provider.lower()}-pending')); c.commit()
    except sqlite3.IntegrityError:
        c.close(); return jsonify({'error':'This service number already exists for the selected Provider and Service Type.'}),409
    c.close(); return jsonify({'ok':True,'id':new_id}),201


@app.route('/api/accounts/<int:account_id>', methods=['PUT','DELETE'])
def manage_account(account_id):
    c=conn(); old=c.execute('SELECT * FROM we_accounts WHERE id=?',(account_id,)).fetchone()
    if not old: c.close(); return jsonify({'error':'Account not found.'}),404
    if request.method=='DELETE':
        c.execute('DELETE FROM we_accounts WHERE id=?',(account_id,)); c.commit(); c.close(); return jsonify({'ok':True})
    data=request.get_json(silent=True) or {}; raw_number=str(data.get('account_number') or data.get('login_username') or '').strip()
    if not raw_number: c.close(); return jsonify({'error':'Service/login number is required.'}),400
    new_provider=str(data.get('provider',old['provider'] or 'WE')).strip() or 'WE'; new_portal=str(data.get('portal_service_type',old['portal_service_type'] or 'Internet')).strip() or 'Internet'; number=account_storage_number(raw_number,new_provider,new_portal); dup=c.execute('SELECT id FROM we_accounts WHERE lower(provider)=lower(?) AND account_number=? AND lower(portal_service_type)=lower(?) AND id<>?',(new_provider,number,new_portal,account_id)).fetchone()
    if dup: c.close(); return jsonify({'error':'This service number already exists for the selected Provider and Service Type.'}),409
    pw=old['password_enc']; new_pw=str(data.get('password',''))
    if new_pw: pw=encrypt_secret(new_pw)
    old_number=old['account_number']; now=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute('''UPDATE we_accounts SET provider=?,name=?,account_number=?,site=?,legal_entity=?,login_username=?,portal_service_type=?,password_enc=?,enabled=?,updated_at=? WHERE id=?''',(str(data.get('provider',old['provider'] or 'WE')).strip() or 'WE',str(data.get('name','')).strip(),number,str(data.get('site','')).strip(),str(data.get('legal_entity','')).strip(),str(data.get('login_username','')).strip() or raw_number,str(data.get('portal_service_type','Internet')).strip() or 'Internet',pw,1 if data.get('enabled',1) else 0,now,account_id))
    if old_number != number: c.execute('UPDATE lines SET account_number=? WHERE account_number=? AND lower(provider)=lower(?) AND lower(service_type)=lower(?)',(number,old_number,old['provider'] or 'WE','LandLine' if str(old['portal_service_type'] or '').lower()=='landline' else ('Home Wireless' if str(old['provider'] or '').lower()=='orange' else 'Internet')))
    svc_old='LandLine' if str(old['portal_service_type'] or '').lower()=='landline' else ('Home Wireless' if str(old['provider'] or '').lower()=='orange' else 'Internet')
    c.execute('UPDATE lines SET provider=?, site=?, legal_entity=?, landline=? WHERE account_number=? AND lower(provider)=lower(?) AND lower(service_type)=lower(?)',(new_provider,str(data.get('site','')).strip(),str(data.get('legal_entity','')).strip(),str(data.get('login_username','')).strip() or number,number,old['provider'] or 'WE',svc_old))
    c.commit(); c.close(); return jsonify({'ok':True})


def ensure_account_placeholder(account):
    """Keep a visible line row before remote provider mapping succeeds."""
    number = str(account.get('account_number') or '').strip()
    login = str(account.get('login_username') or '').strip()
    site = str(account.get('site') or '').strip()
    legal_entity = str(account.get('legal_entity') or '').strip()
    if not number or not login:
        return
    c = conn()
    try:
        existing = c.execute('SELECT id FROM lines WHERE account_number=? AND landline=? AND lower(provider)=lower(?) AND lower(service_type)=lower(?)',(number,login,str(account.get('provider') or 'WE'),'LandLine' if str(account.get('portal_service_type') or '').lower()=='landline' else ('Home Wireless' if str(account.get('provider') or '').lower()=='orange' else 'Internet'))).fetchone()
        if not existing:
            now=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            sql = 'INSERT INTO lines(site,legal_entity,account_number,landline,provider,service_type,package,speed_mbps,quota_gb,used_gb,usage_pct,remaining_gb,bill_egp,renewal_cost_egp,outstanding_egp,due_date,renewal_date,status,last_sync,source) VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)'
            prov=str(account.get('provider') or 'WE'); portal_type=str(account.get('portal_service_type') or 'Internet'); c.execute(sql,(site,legal_entity,number,login,prov,'Home Wireless' if prov.lower()=='orange' else ('LandLine' if portal_type.lower()=='landline' else 'Internet'),'',0,0,0,0,0,0,0,0,'','','Pending Sync',now,f"{prov.lower()}-pending"))
            c.commit()
    finally:
        c.close()


def build_notifications(rows):
    notes=[]
    for x in rows:
        line=str(x.get('landline') or '').strip() or 'Unknown line'
        site=str(x.get('site') or '').strip(); account=str(x.get('account_number') or '').strip()
        rc=nfloat(x.get('renewal_cost_egp')); bal=nfloat(x.get('outstanding_egp'))
        if rc > 0 and bal > 0 and rc > bal:
            diff=rc-bal
            notes.append({'key':f'charge:{account}:{line}','type':'need_charge','severity':'warning','title':'Need to Charge Line','landline':line,'site':site,'account_number':account,'message':f'Renewal Cost {rc:,.2f} EGP is greater than Current Balance {bal:,.2f} EGP. Required charge: {diff:,.2f} EGP.'})
        if str(x.get('status') or '').strip().lower() == 'high usage':
            notes.append({'key':f'renew:{account}:{line}','type':'need_renew','severity':'critical','title':'Need To Renew','landline':line,'site':site,'account_number':account,'message':f'Line {line} is in High Usage status ({nfloat(x.get("usage_pct")):.1f}% used).'})
    return notes


def sync_notification_state(rows):
    alerts=build_notifications(rows); active_keys={a['key'] for a in alerts}; now=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c=conn(); existing={r['alert_key']:dict(r) for r in c.execute('SELECT * FROM notification_state')}
    for a in alerts:
        prev=existing.get(a['key']); reactivated=not prev or not prev.get('active')
        read_at=None if reactivated else (prev.get('read_at') if prev else None)
        last_sent=prev.get('last_sent') if prev else None
        c.execute('''INSERT INTO notification_state(alert_key,active,last_sent,updated_at,read_at) VALUES(?,?,?,?,?)
                     ON CONFLICT(alert_key) DO UPDATE SET active=excluded.active,updated_at=excluded.updated_at,read_at=excluded.read_at''',
                  (a['key'],1,last_sent,now,read_at))
    for key,prev in existing.items():
        if key not in active_keys and prev.get('active'):
            c.execute('UPDATE notification_state SET active=0,updated_at=? WHERE alert_key=?',(now,key))
    c.commit()
    states={r['alert_key']:dict(r) for r in c.execute('SELECT * FROM notification_state WHERE active=1')}
    c.close()
    out=[]
    for a in alerts:
        st=states.get(a['key'],{}); item=dict(a); item['read']=bool(st.get('read_at')); item['read_at']=st.get('read_at') or ''; out.append(item)
    return out

@app.route('/api/notifications/read', methods=['POST'])
def read_notification():
    d=request.get_json(silent=True) or {}; now=datetime.now().strftime('%Y-%m-%d %H:%M:%S'); c=conn()
    if d.get('all'):
        c.execute('UPDATE notification_state SET read_at=? WHERE active=1',(now,))
    else:
        key=str(d.get('key') or '').strip()
        if not key:
            c.close(); return jsonify({'error':'Notification key is required.'}),400
        c.execute('UPDATE notification_state SET read_at=? WHERE alert_key=? AND active=1',(now,key))
    c.commit(); c.close(); return jsonify({'ok':True})



def backup_database(force=False):
    """Create one consistent SQLite backup per day; optionally force a timestamped backup."""
    if not os.path.exists(DB):
        return None
    os.makedirs(BACKUP_DIR, exist_ok=True)
    day = datetime.now().strftime('%Y-%m-%d')
    name = f"we_dashboard_{day}.db" if not force else f"we_dashboard_{datetime.now().strftime('%Y-%m-%d_%H%M%S')}.db"
    target = os.path.join(BACKUP_DIR, name)
    if os.path.exists(target) and not force:
        return target
    src = sqlite3.connect(DB)
    dst = sqlite3.connect(target)
    try:
        src.backup(dst)
    finally:
        dst.close(); src.close()
    # Keep the encryption key beside backups because encrypted WE credentials depend on it.
    if os.path.exists(KEY_FILE):
        try:
            shutil.copy2(KEY_FILE, os.path.join(BACKUP_DIR, '.we_secret.key'))
        except Exception:
            pass
    retention = max(1, int(os.getenv('BACKUP_RETENTION_DAYS','30')))
    cutoff = time.time() - retention * 86400
    for f in Path(BACKUP_DIR).glob('we_dashboard_*.db'):
        try:
            if f.stat().st_mtime < cutoff:
                f.unlink()
        except Exception:
            pass
    return target


def backup_loop():
    # Check hourly, but create at most one normal backup per calendar day.
    time.sleep(10)
    while True:
        try:
            path = backup_database(False)
            if path:
                print('Database backup ready:', path)
        except Exception as e:
            print('Database backup failed:', e)
        time.sleep(3600)


@app.route('/api/system/storage', methods=['GET'])
def api_system_storage():
    backups = sorted([p.name for p in Path(BACKUP_DIR).glob('we_dashboard_*.db')], reverse=True)
    return jsonify({'data_dir':DATA_DIR,'database':DB,'sessions_dir':SESSIONS_DIR,'backup_dir':BACKUP_DIR,'latest_backups':backups[:10]})


@app.route('/api/backup', methods=['POST'])
def api_backup_now():
    try:
        path=backup_database(True)
        return jsonify({'ok':True,'path':path})
    except Exception as e:
        return jsonify({'ok':False,'error':str(e)}),500


def fetch_provider_lines(account, interactive=False):
    provider=(account.get('provider') or 'WE').strip().lower()
    if provider == 'orange':
        from orange_connector import fetch_orange_lines
        return fetch_orange_lines(account, interactive=interactive)
    from we_connector import fetch_we_lines
    return fetch_we_lines(account, interactive=interactive)

def sync_all_accounts(interactive=False):
    if not SYNC_LOCK.acquire(blocking=False):
        return {'ok':False,'error':'A sync is already running.','updated':0,'failed':0}
    started=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    SYNC_STATE.update({'running':True,'last_started':started,'last_error':'','updated':0,'failed':0})
    updated=0; failed=0; messages=[]
    try:
        c=conn(); accounts=[dict(r) for r in c.execute("SELECT * FROM we_accounts WHERE enabled=1 AND lower(coalesce(portal_service_type,'Internet')) <> 'landline'")]; c.close()
        for a in accounts:
            a['password']=decrypt_secret(a.pop('password_enc',None))
            ensure_account_placeholder(a)
            if not a.get('login_username') or not a.get('password'):
                failed += 1; messages.append(f"{a.get('account_number')}: missing credentials"); continue
            try:
                data=fetch_provider_lines(a, interactive=interactive)
                if data:
                    upsert_synced_lines(a['account_number'],data); updated += len(data)
                else:
                    failed += 1; messages.append(f"{a.get('account_number')}: no data returned")
            except Exception as e:
                failed += 1; messages.append(f"{a.get('account_number')}: {e}")
                # Authentication/verification failures must never overwrite the business
                # status of the line. Keep the last successful provider data/status intact and
                # expose the authentication issue only through sync status/notifications.
                print(f"Internet sync failed for {a.get('account_number')}: {e}")
        return {'ok': failed == 0, 'updated':updated, 'failed':failed, 'messages':messages}
    finally:
        finished=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        msg='; '.join(messages)[:3000]
        SYNC_STATE.update({'running':False,'last_finished':finished,'last_error':msg,'updated':updated,'failed':failed})
        try:
            c=conn(); c.execute('INSERT INTO sync_log(started_at,finished_at,updated,failed,message) VALUES(?,?,?,?,?)',(started,finished,updated,failed,msg)); c.commit(); c.close()
        except Exception:
            pass
        SYNC_LOCK.release()




def sync_landline_accounts(interactive=False):
    if not SYNC_LOCK.acquire(blocking=False):
        return {'ok':False,'error':'A sync is already running.','updated':0,'failed':0}
    started=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    SYNC_STATE.update({'running':True,'last_started':started,'last_error':'','updated':0,'failed':0,'target_line':'WE Land Lines'})
    updated=0; failed=0; messages=[]
    try:
        c=conn(); accounts=[dict(r) for r in c.execute("SELECT * FROM we_accounts WHERE enabled=1 AND lower(provider)='we' AND lower(portal_service_type)='landline'")]; c.close()
        for a in accounts:
            a['password']=decrypt_secret(a.pop('password_enc',None))
            ensure_account_placeholder(a)
            if not a.get('login_username') or not a.get('password'):
                failed += 1; messages.append(f"{a.get('account_number')}: missing credentials"); continue
            try:
                data=fetch_provider_lines(a, interactive=interactive)
                if data:
                    upsert_synced_lines(a['account_number'],data); updated += len(data)
                else:
                    failed += 1; messages.append(f"{a.get('account_number')}: no data returned")
            except Exception as e:
                failed += 1; messages.append(f"{a.get('account_number')}: {e}"); print(f"WE LandLine sync failed for {a.get('account_number')}: {e}")
        return {'ok':failed==0,'updated':updated,'failed':failed,'messages':messages}
    finally:
        finished=datetime.now().strftime('%Y-%m-%d %H:%M:%S'); msg='; '.join(messages)[:3000]
        SYNC_STATE.update({'running':False,'last_finished':finished,'last_error':msg,'updated':updated,'failed':failed,'target_line':'WE Land Lines'})
        try:
            c=conn(); c.execute('INSERT INTO sync_log(started_at,finished_at,updated,failed,message) VALUES(?,?,?,?,?)',(started,finished,updated,failed,msg)); c.commit(); c.close()
        except Exception: pass
        SYNC_LOCK.release()

@app.route('/api/sync/landlines', methods=['POST'])
def api_sync_landlines():
    if SYNC_STATE.get('running'):
        return jsonify({'ok':False,'error':'A sync is already running.'}),409
    threading.Thread(target=sync_landline_accounts,kwargs={'interactive':True},daemon=True).start()
    return jsonify({'ok':True,'started':True}),202

def sync_one_line(line_id, interactive=True):
    if not SYNC_LOCK.acquire(blocking=False):
        return {'ok':False,'error':'A sync is already running.','updated':0,'failed':0}
    started=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    updated=0; failed=0; messages=[]
    target_landline=''
    try:
        c=conn()
        line=c.execute('SELECT * FROM lines WHERE id=?',(line_id,)).fetchone()
        if not line:
            c.close(); failed=1; messages.append('Line not found'); return {'ok':False,'updated':0,'failed':1,'messages':messages}
        line=dict(line); target_landline=str(line.get('landline') or '').strip()
        account=c.execute('SELECT * FROM we_accounts WHERE account_number=? AND lower(provider)=lower(?) AND lower(portal_service_type)=lower(?) AND enabled=1',(line.get('account_number'),line.get('provider') or 'WE',portal_type_for_line(line.get('service_type'),line.get('provider')))).fetchone()
        c.close()
        if not account:
            failed=1; messages.append(f'{target_landline}: linked internet account not found or disabled'); return {'ok':False,'updated':0,'failed':1,'messages':messages}
        a=dict(account); a['password']=decrypt_secret(a.pop('password_enc',None))
        if not a.get('login_username') or not a.get('password'):
            failed=1; messages.append(f'{target_landline}: missing credentials'); return {'ok':False,'updated':0,'failed':1,'messages':messages}
        SYNC_STATE.update({'running':True,'last_started':started,'last_error':'','updated':0,'failed':0,'target_line':target_landline})
        try:
            data=fetch_provider_lines(a, interactive=interactive) or []
            # Only update the selected line even if the provider account returns more than one service.
            matched=[item for item in data if str(item.get('landline') or '').strip()==target_landline]
            if not matched and len(data)==1:
                # Some WE responses omit the service number; when exactly one record is returned,
                # bind it to the selected line explicitly.
                item=dict(data[0]); item['landline']=target_landline; matched=[item]
            if matched:
                upsert_synced_lines(a['account_number'], matched)
                updated=len(matched)
            else:
                failed=1; messages.append(f'{target_landline}: Provider returned no matching data for this line')
        except Exception as e:
            failed=1; messages.append(f'{target_landline}: {e}'); print(f'Internet line sync failed for {target_landline}: {e}')
        return {'ok':failed==0,'updated':updated,'failed':failed,'messages':messages}
    finally:
        finished=datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        msg='; '.join(messages)[:3000]
        SYNC_STATE.update({'running':False,'last_finished':finished,'last_error':msg,'updated':updated,'failed':failed,'target_line':target_landline})
        try:
            c=conn(); c.execute('INSERT INTO sync_log(started_at,finished_at,updated,failed,message) VALUES(?,?,?,?,?)',(started,finished,updated,failed,msg)); c.commit(); c.close()
        except Exception: pass
        SYNC_LOCK.release()


@app.route('/api/sync/line/<int:line_id>', methods=['POST'])
def api_sync_line(line_id):
    if SYNC_STATE.get('running'):
        return jsonify({'ok':False,'error':'A sync is already running.'}),409
    # Validate before starting a background thread so obvious errors return immediately.
    c=conn(); exists=c.execute('SELECT id FROM lines WHERE id=?',(line_id,)).fetchone(); c.close()
    if not exists:
        return jsonify({'ok':False,'error':'Line not found.'}),404
    threading.Thread(target=sync_one_line,args=(line_id,),kwargs={'interactive':True},daemon=True).start()
    return jsonify({'ok':True,'message':'Sync started for selected line only.','line_id':line_id}),202


@app.route('/api/sync', methods=['POST'])
def api_sync():
    if SYNC_STATE.get('running'):
        return jsonify({'ok':False,'error':'A sync is already running.'}), 409
    threading.Thread(target=sync_all_accounts, kwargs={'interactive':True}, daemon=True).start()
    return jsonify({'ok':True,'message':'Internet Lines sync started in background.'}), 202


@app.route('/api/background/stop', methods=['POST'])
def api_background_stop():
    """Stop this dashboard process after returning a response to the browser."""
    def _shutdown():
        time.sleep(0.8)
        os._exit(0)
    threading.Thread(target=_shutdown, daemon=True).start()
    return jsonify({'ok': True, 'message': 'Internet Dashboard background service is stopping.'}), 202


@app.route('/api/sync/status', methods=['GET'])
def api_sync_status():
    return jsonify(SYNC_STATE)


def sync_loop():
    mins=max(1,int(os.getenv('SYNC_INTERVAL_MINUTES','60')))
    # The scheduled sync is provider-neutral: every enabled WE and Orange Home Wireless
    # account is processed headlessly. Orange reuses its saved session and, when needed,
    # performs a normal credential login in the background. CAPTCHA still requires a
    # one-time manual sync because it cannot be bypassed automatically.
    while True:
        time.sleep(mins*60)
        try:
            internet_result = sync_all_accounts(interactive=False)
            print(f"Background Internet sync completed: updated={internet_result.get('updated',0)}, failed={internet_result.get('failed',0)}")
            landline_result = sync_landline_accounts(interactive=False)
            print(f"Background LandLine sync completed: updated={landline_result.get('updated',0)}, failed={landline_result.get('failed',0)}")
        except Exception as e:
            print('Scheduled sync loop failed:',e)


def normalize_legacy_sync_statuses():
    """Repair statuses written by older builds that mixed auth state with line state."""
    try:
        c=conn()
        rows=c.execute("SELECT id,status,package,quota_gb,last_sync,source FROM lines WHERE status IN ('Verification Required','Pending Sync','Partial Sync')").fetchall()
        for r in rows:
            # If we already have real provider data from a successful sync, restore Active.
            # Otherwise keep Pending Sync rather than presenting auth state as line state.
            has_data = bool((r['package'] or '').strip()) or float(r['quota_gb'] or 0) > 0 or bool(r['last_sync'])
            new_status = 'Active' if has_data else 'Pending Sync'
            c.execute('UPDATE lines SET status=? WHERE id=?',(new_status,r['id']))
        c.commit(); c.close()
    except Exception as e:
        print('Status normalization skipped:', e)

if __name__=='__main__':
    init_db()
    normalize_legacy_sync_statuses()
    print('\n' + '='*72)
    print('Internet Dashboard build:', APP_BUILD)
    print('Operational Dashboard: http://127.0.0.1:5000/')
    print('Visual Dashboard:      http://127.0.0.1:5000/visual')
    print('WE Land Lines:         http://127.0.0.1:5000/landlines')
    print('Version check:         http://127.0.0.1:5000/version')
    print('Persistent data folder:', DATA_DIR)
    print('Database file:         ', DB)
    print('Backups folder:        ', BACKUP_DIR)
    print('='*72 + '\n')
    try:
        backup_database(False)
    except Exception as e:
        print('Startup backup warning:', e)
    threading.Thread(target=sync_loop,daemon=True).start()
    threading.Thread(target=backup_loop,daemon=True).start()
    app.run(host='0.0.0.0',port=5000,debug=False)
