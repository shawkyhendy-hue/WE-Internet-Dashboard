# Internet Dashboard v5.10 Extended

This package is based on **Internet Dashboard v5.10 — Branch** and adds only these operational features:

- Automatic WE synchronization every 60 minutes while the background process is running.
- Windows Auto Start/background launch so `run_dashboard.bat` is not required for daily use.
- **Stop Background** control from the Operational Dashboard plus `STOP_BACKGROUND.bat`.
- Export current line data to **Excel (.xlsx)** or **CSV**, respecting the active dashboard filters.
- No **Stop Sync** feature is included in this build.

The existing persistent database under `%USERPROFILE%\WE-Dashboard-Data` is reused automatically.

---

# WE Internet Multi-Line Dashboard - Automatic Live Sync (Python 3.13)

This edition logs in to **My WE** automatically for every enabled account, selects **Internet**, listens to authenticated JSON/XHR responses, extracts line/package/usage/billing fields, and writes them to the local SQLite dashboard cache.

## What is automatic now
- Multiple WE accounts and multiple Internet lines.
- Login flow: Service Number -> Service Type = Internet -> Password -> Login.
- Background sync every `SYNC_INTERVAL_MINUTES` (default 60 minutes).
- **Sync WE Now** button for an immediate live pull.
- Dashboard refresh every 60 seconds by default.
- Automatic upsert of lines into SQLite.
- Data source appears as `we` for synchronized rows.
- Saved WE passwords are encrypted locally.
- No sample lines are inserted by default (`SEED_SAMPLE_DATA=0`).

## Install on Windows / Python 3.13
1. Extract the ZIP fully.
2. Install Python 3.13 x64 if it is not installed.
3. Run `install_python313.bat`.
4. Run `run_dashboard.bat`.
5. Open `http://127.0.0.1:5000`.

## Add a WE account
In the dashboard choose **+ Add WE Account** and enter:
- Account Name (your own label)
- WE Account Number (internal dashboard identifier; it can be the same as Service Number)
- Site / Branch
- WE Service Number used on My WE
- Service Type = Internet
- My WE password
- Automatic Sync = Enabled

Then press **Sync WE Now**. The background job will continue automatically at the configured interval.

## First live test
For the first account, edit `.env` and use:

```text
WE_BROWSER_HEADLESS=0
SYNC_INTERVAL_MINUTES=60
WE_DIAGNOSTIC_CAPTURE=1
```

Restart `run_dashboard.bat`, then click **Sync WE Now**. Chromium will show the WE login flow. After the data is confirmed, change `WE_BROWSER_HEADLESS=1`.

## How extraction works
The connector does not hard-code an undocumented WE private endpoint. It captures authenticated JSON/XHR responses produced by the portal itself and maps common response fields such as service number, package, quota, consumed/remaining data, bill/outstanding amount, renewal date, due date, status and speed. It also has a conservative visible-page fallback.

Because WE can change its internal JSON field names or UI at any time, a first live run may leave some columns blank. If that happens, enable `WE_DIAGNOSTIC_CAPTURE=1`. A diagnostics file is created in `diagnostics/` containing response URLs, JSON shapes/keys and a page-text preview (not the raw authenticated JSON). That file can be used to tune the field mapping without sharing your password.

## Important security notes
- Do not send your WE password in chat.
- Keep `.we_secret.key`, `.env`, `we_dashboard.db` and the `diagnostics` folder private.
- This connector performs read-only navigation and intentionally avoids recharge/payment/renewal actions.


## Fix 2026-09-12 — WE Service Type / sync UI
- Fixed WE Ant Design Service Type selector (`login_input_type_01`) so `Internet` is selected from the actual `.ant-select-selector` / option node.
- Manual `Sync WE Now` now starts in the background and the dashboard polls `/api/sync/status`; the button no longer waits on the browser automation HTTP request.
- If sync fails, the exact connector error is shown after the background sync finishes.

## Mapping diagnostics build
If login succeeds but the dashboard still shows zero/blank package, quota, usage or bill values, enable:

WE_DIAGNOSTIC_CAPTURE=1

Then press **Sync WE Now** once. This build will NOT report a successful sync unless at least one real usage/package/billing field is extracted. It writes a sanitized mapping file under `diagnostics/we_mapping_*.json` containing endpoint paths and candidate field names. Passwords, auth tokens, cookies and long service numbers are redacted.

Upload that JSON file to finish mapping the current WE portal fields precisely.

## Fix 2026-09-12 — forced diagnostics and pending-line preservation
- If login succeeds but WE field mapping fails, a sanitized diagnostics JSON is now created **automatically**, even when `WE_DIAGNOSTIC_CAPTURE=0`.
- The CMD window prints the exact diagnostics path: `WE mapping diagnostic saved to: ...`.
- XHR/fetch responses are inspected even when WE does not label them `application/json`.
- A WE account now keeps a visible `Pending Sync` line row using its Service Number, so a mapping failure does not make the line disappear from the dashboard.
- When moving to this build and you want to preserve existing accounts, copy `we_dashboard.db` and `.we_secret.key` from the old project folder into this project folder before starting it.


## Interactive WE verification / CAPTCHA
If WE shows **Verification / Please enter the code below**, the connector intentionally does not bypass it.
Chrome remains open for up to `WE_CAPTCHA_TIMEOUT_SECONDS` (default 240 seconds). Enter the code manually in the WE window and submit it. The connector then detects the authenticated Account Overview and continues automatically to Usage and Bills & Payments capture. Keep `WE_BROWSER_HEADLESS=0` while WE requires this verification.

## Exact WE visible-page mapping (Sep 2026)
This build prioritizes the authenticated WE page labels over generic JSON balances:
- Current Plan -> Package
- Main Quota `remaining GB / total GB` -> Remaining and Quota
- Circular widget `N Used` -> Used
- Renewal Cost -> Monthly Bill
- Renewal Date -> Renewal and dashboard Due date
- Usage % -> calculated as Used / Quota * 100

## 2026-09-12 field corrections
- Renewal Cost is stored separately from Bill.
- Current Balance from WE Home is mapped to Outstanding.
- Due displays Remaining Days (for example `9 Remaining Days`).
- Renewal Date remains the actual date (for example `21-09-2026`).
- Sync visits Home, Account Overview, Usage, and Bills & Payments before merging values.


## Hourly background sync and smart CAPTCHA behavior
- Automatic synchronization now runs every **60 minutes** by default.
- Scheduled syncs always run **headless/in the background** and do not open Chrome.
- Authenticated browser session state is saved under `sessions/` and reused on future syncs, reducing repeated logins and CAPTCHA prompts.
- If the saved WE session expires and WE requires Verification/CAPTCHA, the hourly job records **Verification Required** and does not open a browser.
- Press **Sync WE Now** when that status appears. The connector first retries silently; only if WE still requests verification will visible Chrome open so you can enter the code manually.
- After successful verification the new session is saved, and subsequent hourly syncs return to background mode automatically.

Recommended production settings:
```text
SYNC_INTERVAL_MINUTES=60
WE_BROWSER_HEADLESS=1
WE_USE_INSTALLED_CHROME=1
WE_DIAGNOSTIC_CAPTURE=0
WE_CAPTCHA_TIMEOUT_SECONDS=240
```

## Visual Dashboard
This build adds a second graphical dashboard inspired by the provided reference image.

- Operational dashboard: http://127.0.0.1:5000/
- Visual dashboard: http://127.0.0.1:5000/visual

The visual dashboard shows one graphical card per Internet line with:
- Status, site, package, renewal information, renewal cost, current balance and last update
- Circular usage gauge
- Total package, used, remaining and current balance cards
- Summary KPIs and filters
- Manual Sync From WE button

The hourly background synchronization behavior from the previous build is preserved.

## Persistent data folder and automatic backups (v4)

This build stores operational data **outside the application/version folder** so upgrading the ZIP no longer requires re-adding WE lines or accounts.

Default Windows location:

```text
%USERPROFILE%\WE-Dashboard-Data\
    we_dashboard.db
    .we_secret.key
    sessions\
    diagnostics\
    backups\
```

What remains persistent there:
- WE accounts and line records
- Synchronized usage / renewal values
- Saved My WE browser sessions
- Diagnostics and automatic database backups

### One-time migration from your current/old version

Run:

```text
MIGRATE_OLD_DATA.bat
```

Paste the full path of the previous dashboard folder. The script copies `we_dashboard.db`, `.we_secret.key`, and saved sessions into the persistent folder. Do this once only. Future ZIP versions can use the same data automatically.

You can also place the old `we_dashboard.db` and `.we_secret.key` beside `app.py` before the first launch of this version; the app detects them and performs a one-time migration automatically if the persistent folder is still empty.

### Automatic backup

The application creates one consistent SQLite backup per day under:

```text
%USERPROFILE%\WE-Dashboard-Data\backups
```

Default retention is 30 days and can be changed in `.env`:

```text
BACKUP_RETENTION_DAYS=30
```

For an immediate manual backup while the dashboard is running, use `BACKUP_NOW.bat`.

Use `OPEN_DATA_FOLDER.bat` to open the persistent storage folder quickly.

- Added a bell icon to both Operational and Visual dashboards.
- The red badge shows the number of unread active notifications.
- Click a notification to mark it as read; the badge count decreases immediately.
- Use **Mark all read** to clear the unread badge while keeping active alerts visible in the notification list.
- If an alert condition clears and later becomes active again, it becomes unread again automatically.
- Existing alert rules remain unchanged:
  - Renewal Cost > Current Balance / Outstanding -> **Need to Charge Line** (with line number and required charge amount).
  - Status = High Usage -> **Need To Renew** (with line number and usage percentage).
- Browser notification permission is also available from **Setting**.

## v5.7 - Internet Dashboard / no email configuration

- Main title changed to **Internet Dashboard**.
- The **Setting** menu was removed from both dashboards.
- Email/SMTP delivery was removed from the application.
- On first startup, the legacy `email_settings` table is deleted from the persistent SQLite database and from retained SQLite backups.
- The in-app bell Notification Center remains enabled.
- Visual Dashboard remains read-only and uses the shared Operational Dashboard data.


## v5.8 - Legal Entity
- Added Legal Entity to line records.
- Added Legal Entity filter to Operational and Visual dashboards.
- Added Legal Entity to line details and Visual cards.
- Added Default Legal Entity to Add/Edit WE Account; newly created placeholder lines inherit it.
- Existing databases are migrated automatically with no data loss.
- Existing line Legal Entity values are preserved during WE sync when the connector does not provide this field.

## v5.9 - Account Name
- Added Account Name as a column in the Operational Dashboard.
- Added Account Name filter to Operational and Visual dashboards.
- Account Name is resolved from the linked WE Account, so changing an account name updates all linked lines automatically.
- Added Account Name to line details, Visual cards, and search.

## v5.11 Always-On Auto Sync
- Run `INSTALL_AUTOSTART.bat` once. It creates a Windows Startup shortcut that launches `app.py` with `pythonw.exe`, so no CMD window and no need to run `run_dashboard.bat` after each Windows sign-in.
- The server stays active in the background and performs WE sync every 60 minutes (`SYNC_INTERVAL_MINUTES=60`).
- Opening the main Operational Dashboard (`/`) does **not** trigger a WE sync. The first automatic sync is handled only by the hourly background scheduler; manual Sync buttons remain available.
- The Visual Dashboard remains read-only and never starts a WE sync. It continues to display the same shared database data.
- If WE expires the authenticated session and requires CAPTCHA/Verification, background sync cannot bypass it. Use the main dashboard manual sync once to complete verification; saved sessions can then be reused by subsequent background syncs.
- `REMOVE_AUTOSTART.bat` removes the Windows Startup shortcut.


## v5.12 - Export Lines
- Export all or currently filtered line data to Excel (.xlsx) or CSV.
- Export respects Search, Site, Legal Entity, Branch, and Status filters.
- Excel export includes formatting, filters, frozen header, and timestamp.

## v5.13 - Background control

The Operational Dashboard now includes a **Stop Background** button. It stops the currently running hidden dashboard server and the hourly Auto Sync process. This is a temporary stop and does not remove Windows Auto Start.

You can also run `STOP_BACKGROUND.bat` if the browser is not open. To start the dashboard again without opening a CMD window, run `START_BACKGROUND_NOW.bat`. To permanently disable startup with Windows, use `REMOVE_AUTOSTART.bat`.

## Reports tab
This build adds `/reports` with four report views: Lines, Usage, Cost, and Renewals. Reports support Search, Legal Entity, Branch, Site, and Status filters. Use **Export Report Excel** to export the currently selected report and active filters. The workbook includes a Summary sheet plus the selected report sheet.

## Orange Egypt Home Wireless support
This build supports two providers from the same dashboard: WE and Orange.

Use **Add Internet Account** and choose **Orange**. Enter the Orange Home Wireless line/login number, Branch, Legal Entity and Site. The password is optional because Orange Home Wireless may authorize account pages from the DSL service context/session rather than a normal username/password flow.

Orange data is displayed in the same Operational, Visual and Reports views and can be filtered by **Provider**. Sync, notifications, reports and Excel export use the shared lines database.

### First Orange sync
Orange's Home Wireless site may require the browser to be connected through the relevant Orange DSL service/location. If Sync reports that the portal requires the DSL service context, run a manual Sync from a PC/network using that Orange DSL line. The connector saves the authorized browser session for later background attempts.

If Orange opens successfully but fields are not yet recognized, a safe mapping file named `orange_mapping_*.json` is written in `%USERPROFILE%\WE-Dashboard-Data\diagnostics`. It does not intentionally store passwords/tokens. That file can be used to tune field mapping for the exact Orange account page.


## Orange Home Wireless
Orange Home Wireless is treated as a mobile/data SIM line. Login uses https://www.orange.eg/ar/myaccount/login?ReturnUrl=%2Far%2F with the Orange mobile number and Orange account password. If Orange requests CAPTCHA, run a manual per-line Sync and complete it in the opened Chrome window; the saved session is reused by later syncs.


## v5.17 - Orange Home Wireless exact mapping and visual design

Orange Home Wireless is mapped from `https://www.orange.eg/ar/myaccount/internet` using the visible page text.

- Package: `Home Wireless N GB`
- Renewal Cost: amount after the package title (`EGP ...`)
- Quota: N GB from the package title
- Remaining: Arabic `السعه المتبقية:...MBs`, converted using MB / 1024
- Used: Quota - Remaining
- Usage %: Used / Quota * 100
- Renewal Date: Arabic `تاريخ التجديد` when present

The Visual Dashboard now uses the new WE + Orange design and remains read-only/shared-data as previously requested.

## v5.18 - Orange direct Internet-page mapping
Orange Home Wireless now uses the authenticated page below as the primary mapping source:
https://www.orange.eg/ar/myaccount/internet

After login the connector waits for the Home Wireless usage panel, then maps the visible DOM text directly. For Orange only:
- Package: Home Wireless <quota> GB
- Renewal Cost: from the package header
- Quota: from the package header
- Remaining: from `السعه المتبقية` in MB, converted using MB / 1024
- Used: Quota - Remaining
- Usage %: Used / Quota * 100
- Renewal Date: from `تاريخ التجديد`

Other Orange account pages are no longer mixed into the Home Wireless mapping.

## v5.19 - Unified UI and Orange Due
- Orange Home Wireless Due is calculated automatically as Renewal Date minus today's date and stored as `N Remaining Days`.
- Operational Dashboard and Reports now use the same navy sidebar / blue header visual language as the Visual Dashboard.
- Visual Dashboard no longer shows Orange mapping explanations or Calculation Details; it shows operational values only.
- Visual Dashboard now shows Due in the line summary, details, and line table.

## v5.20 changes
- Orange Home Wireless is included in the scheduled background sync every `SYNC_INTERVAL_MINUTES` (60 by default).
- Orange background sync reuses the saved session and can refresh the Orange session headlessly with the stored credentials. If Orange requests CAPTCHA, run one manual sync to refresh the saved session.
- Removed the top-header navigation buttons for Operational Dashboard, Visual Dashboard, and Reports. Navigation remains available in the left sidebar.
- Added an active `Refresh View` button to the Visual Dashboard. It refreshes data from the shared local database only and does not start a provider sync.


## v5.21 - Orange silent sync / CAPTCHA-only browser
- Orange Home Wireless Sync now always starts headless, including manual per-line Sync.
- The Orange login page is not shown during normal login or session refresh.
- Chrome becomes visible only when Orange positively requires CAPTCHA/interactive verification.
- Background Sync remains fully headless. If CAPTCHA is required during background Sync, it records the verification requirement without opening a browser; run manual Sync once to complete it.
- WE connector behavior is unchanged.


## v5.25 Visual Dashboard
Removed Legal Entity, Branch, Site, and Renewal Date from each Visual Dashboard line card. Filters remain available.


## v5.26 - Sortable Columns & Quota
- Added click-to-sort to Operational Dashboard table columns (ascending/descending).
- Removed Speed column from Operational Dashboard.
- Added Quota column in the former Speed position.
- Corrected Operational table cell order to match Legal Entity / Site / Branch headers.


## v5.27
- Removed the Stop Background button from Visual Dashboard only.
- Stop Background remains available from Operational Dashboard.

## v5.28 - WE Land Lines
- Added a new **Land Lines** tab in the left navigation.
- Dedicated WE LandLine account form with Portal Service Type fixed to **LandLine**.
- Dedicated **Sync Land Lines** action plus per-line Sync.
- WE connector now preserves `service_type=LandLine` and does not require Internet quota/usage fields for LandLine sync.
- LandLine accounts participate in the existing scheduled/background sync because they are stored as enabled WE accounts.
- Existing WE Internet and Orange Home Wireless mappings are unchanged.

## v5.31 - Strict Internet / LandLine separation
WE LandLine accounts now use an internal namespaced key while keeping the visible/login number unchanged. This allows the same service number to exist once as WE Internet and once as WE LandLine, even when upgrading from databases that had a legacy UNIQUE(account_number) constraint.

## v5.32 changes
- LandLine services are excluded from Visual Dashboard and Reports.
- Land Lines page has its own filtered Excel export.
- Visual Dashboard shows Outstanding instead of Renewal Cost in the fourth KPI box.
- Outstanding box turns red when Renewal Cost is greater than Outstanding.

## v5.33 sync scope update
- The **Sync** button on Internet Lines syncs Internet services only (WE Internet + Orange Home Wireless). WE LandLine accounts are excluded.
- The **Sync** button on Land Lines syncs WE LandLine services only.
- The hourly scheduler runs the Internet and LandLine scopes sequentially so both remain updated without mixing service types.
- Account display names fall back to the service/line number instead of the generic `Internet Account` label when no Branch/name is set.
