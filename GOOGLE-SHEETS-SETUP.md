# Digital-Txn — restore Google Spreadsheet and Apps Script

This adds a new backend for the Digital-Txn dashboard. It cannot recover rows that existed only in a deleted Google Sheet. Before creating a replacement, check Google Drive → Trash and restore the old spreadsheet if it is there.

## 1. Create the spreadsheet
1. Open https://sheets.google.com and create a blank spreadsheet named **Digital-Txn — Nagpur Region**.
2. In that spreadsheet choose **Extensions → Apps Script**.
3. Replace the contents of `Code.gs` with the code in this repository's `Code.gs`.
4. Save. In the function dropdown select `setupSpreadsheet`, click **Run**, and authorize the requested Google permissions.
5. Return to the spreadsheet. It should contain **Digital Txn Data** and **Sync Status** tabs.

## 2. Deploy the web app
1. In Apps Script click **Deploy → New deployment**.
2. Select type **Web app**.
3. Set **Execute as: Me**.
4. Set access to **Anyone** if your organisation's Google Workspace policy permits it. Otherwise use the access policy required by your organisation.
5. Click **Deploy**, authorize if prompted, and copy the Web app URL ending in `/exec`.

## 3. Verify
Replace `YOUR_WEB_APP_URL` below with your deployed URL:
- `YOUR_WEB_APP_URL?action=health` — should return JSON with `"ok":true`.
- `YOUR_WEB_APP_URL?action=getRecords` — should return JSON containing `records` and `count`.

The current dashboard posts `{action:"appendRecords", records:[...]}` to the Apps Script URL configured in `index.html`. The new URL must replace `GSHEET_WEB_APP_URL` in `index.html`.

## 4. Automatic download on opening
The existing HTML loads IndexedDB locally, but it does not yet read Google records on every opening. To finish that part, the frontend must call `?action=getRecords` (using JSONP with a `callback` parameter if browser CORS blocks fetch), merge Google records by `txn_date + office_id`, save the merged data into IndexedDB, and refresh filters/KPIs. The backend file alone does not change that frontend behavior.
