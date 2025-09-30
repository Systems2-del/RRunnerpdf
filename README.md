# Dispatch Details – PDF A4 + ≤1MB + Drive Upload (OAuth Refresh Token)

This repo runs a GitHub Action that:
- reads source PDF URLs from **Dispatch Details!I** and invoice numbers from **G**
- normalizes each PDF to **A4** and compresses to **≤ 1 MB**
- uploads to a Drive folder (DEST_FOLDER_ID), sets public “anyone can view”
- writes the viewable link to **L**
- writes a status/log to **M**
- skips rows already processed (L not empty)

## Setup

### A) Add repo **Secrets**
- `GOOGLE_CLIENT_ID`
- `GOOGLE_CLIENT_SECRET`
- `GOOGLE_REFRESH_TOKEN`

(These are from your OAuth app + the refresh token you generated locally.)

### B) Add repo **Variables**
- `SHEET_ID` = `1JPJ5q2vTvybxlsX29r-YXL3mUBH0Piewz2AIJlSfLwc`
- `SHEET_NAME` = `Dispatch Details`
- `DEST_FOLDER_ID` = your Google Drive folder ID where compressed PDFs should be saved
- *(optional)* `START_ROW` = `2`

### C) Share access
Share the Spreadsheet and the Drive folder with the **SAME Google account** you used to obtain the refresh token (Editor on Sheet; Editor/Manager on folder).

### Run
Go to **Actions → “Dispatch Details: Compress & Upload PDFs” → Run workflow**.

