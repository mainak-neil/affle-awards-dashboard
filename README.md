# Affle Awards Dashboard (Streamlit + Google Sheets)

Edit the Google Sheet and the dashboard updates itself. No prompts, no code changes.

## 1. Create the Google Sheet (5 min)
1. Upload `awards_data_template.xlsx` to Google Drive, then **Open with → Google Sheets**
   (then **File → Save as Google Sheets** so it becomes a native sheet).
2. It has three tabs: **How to use**, **Entries** (your data, one row per award × BU) and **Settings** (FX rate, fiscal-year start).
3. Click **Share → General access → Anyone with the link → Viewer**, and copy the link.
   *(Need it private? See "Private sheet" below.)*

## 2. Run it on your laptop (optional, to try it)
```bash
pip install -r requirements.txt
streamlit run app.py
```
Paste the sheet link into the sidebar. Without a link it shows the bundled sample file.

## 3. Put it online for the team (free) — Streamlit Community Cloud
1. Create a GitHub repo and upload everything in this folder (keep the `assets/` and `.streamlit/` folders).
2. Go to https://share.streamlit.io → **Create app** → pick the repo, branch, and `app.py`.
3. In **Advanced settings → Secrets**, paste:
   ```toml
   sheet_url = "https://docs.google.com/spreadsheets/d/YOUR_SHEET_ID/edit"
   ```
4. Deploy. You get a link like `https://affle-awards.streamlit.app` to share.
   You can restrict viewers to specific emails in the app's **Sharing** settings.

## Updating data
- Edit the **Entries** tab. The app re-reads the sheet every 5 minutes, or click **🔄 Refresh data**.
- New award / BU / quarter → just add rows. Every dropdown, chart and total picks it up automatically.
- Quarter comes from **Result Date** (or Submission Date). Indian FY by default (Apr–Jun = Q1);
  change **Fiscal Year Start Month** in Settings to `1` for calendar quarters. Or type a quarter like `Q2 FY27` in the Quarter column.

## Private sheet (no public link)
1. In Google Cloud Console create a project → enable **Google Sheets API** and **Google Drive API**.
2. Create a **service account** → Keys → Add key → JSON.
3. Share the Google Sheet with the service account's `client_email` as **Viewer**.
4. Add the JSON fields under `[gcp_service_account]` in Secrets (see `.streamlit/secrets.toml.example`).

## What's in the app
Sidebar filters: Quarter · Business unit · Status · Award · Currency (US$/₹)
Tabs: Overview · By quarter (BU × quarter matrix) · By BU · By award (cost per win) ·
Pending & upcoming · Case studies · **Build your own view** (pick group-by, split-by, metric, chart) · Data (CSV download)
