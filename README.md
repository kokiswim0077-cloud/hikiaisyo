# Inquiry Voice Form

Japanese voice/photo input web app for generating inquiry Excel files from an existing workbook template.

## What It Does

- Provides an inquiry form at `/` and a quote request/estimate form at `/quote`.
- Accepts Japanese voice or typed text such as:
  `得意先 良栄社、受注日 今日、出荷希望日 明後日、倉庫011、値引き外掛け、値引き率3%、製品SP853A`
- Uses Gemini API when configured, with local parsing fallback for text.
- Uses Gemini Vision for order-form image recognition.
- Suggests customer, delivery destination, and product candidates from the workbook master sheets.
- Handles ambiguous names by showing candidates for user confirmation.
- Generates a new Excel file from `引合書+値引`.
- Saved Excel filenames use creation date, delivery name, and model name:
  `260628_納入先_機種名_引合書.xlsx`.
- Quote estimate filenames use:
  `260628_得意先_機種名_見積書.xlsx`.
- Quote request copy filenames use:
  `260628_得意先_機種名_見積依頼.xlsx`.
- Output files are stored in Google Drive when a local Google Drive sync folder is detected, under `引合書_見積書`.

## Current Business Rules

- If order date is omitted, use today.
- If quantity is omitted, use `1`.
- If production date is omitted, write `在庫` to the production-date cell.
- If production date is provided and shipping date is omitted, shipping date is two business days after production date. Saturdays and Sundays are skipped.
- Warehouse is `011` or `031`.
- Warehouse should be read from the shipping source/remarks, not fixed by customer:
  `011` for OREC head office/Fukuoka warehouse, `031` for OREC Kanto logistics / `オーレック関物`.
- If `033` is read from an old note or OCR result, the app must convert it to `031`.
- Voice input also understands shipping-source words such as `オーレック関物`, `関東物流`, `福岡倉庫`, and `本社`.
- Voice input accepts discount rates such as `3%`, `値引3`, `値引率3`, and `値引き率3パーセント`.
- Discount method is `外掛` or `内掛`.
- Kubota `注文書（出荷指示書）` forms are handled with dedicated overrides:
  customer `61110`, `大網営業所` delivery `61110005`, and model-name product matching such as `RM953X/K` -> `0343-0210`.
- Additional learned image rules:
  `市原営業所` -> delivery `61110004`,
  `木嶋商店` -> customer/delivery `61310`,
  `エルタ` order cards -> customer/delivery `65137`,
  `大竹産業` -> customer/delivery `61323`,
  OREC internal addressees such as `オーレック`, `関東営業所`, or `谷尾` should not become delivery destinations,
  and month-only production/shipping notes are preserved as text in the Excel output.
- Quote form excludes response deadline, freight, payment terms, and FAX status.
- Quote learned rules include:
  `伊藤産業機械` -> customer `61105`,
  `大竹産業` -> customer `61323`,
  and models such as `AM64B`, `RM832A`, `SP853A`, `RCSP540`, `AM65B` resolve through the product master.

## Required Local Files

This handoff repository may include the real Excel template as `template.xlsx` because the user explicitly approved sharing the operational workbook data for handoff.

If `template.xlsx` is not present, put the workbook at:

```text
template.xlsx
```

or set:

```text
INQUIRY_TEMPLATE=C:\secure\path\template.xlsx
```

in `.env`.

Keep repositories containing `template.xlsx` private unless the user explicitly approves publishing the workbook data.

## Google Drive Storage

Install Google Drive for desktop on the app host PC. The app auto-detects common sync folders such as:

```text
G:\マイドライブ
G:\My Drive
%USERPROFILE%\Google Drive
```

When detected, generated Excel files are saved under:

```text
引合書_見積書
```

If auto-detection does not find your folder, run:

```powershell
powershell -ExecutionPolicy Bypass -File .\set_google_drive_output.ps1 "G:\マイドライブ"
```

Then restart the app.

## Setup

```powershell
cd C:\Users\koki0\inquiry_voice_form
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
copy .env.example .env
```

Then set:

```text
GEMINI_API_KEY=...
APP_PASSWORD=...
```

## Run Locally

```powershell
powershell -ExecutionPolicy Bypass -File .\start_local.ps1
```

Open:

```text
http://127.0.0.1:8765
```

## Production Notes

Keep the Flask app bound to `127.0.0.1`. For other PCs, put a reverse proxy, VPN, or Cloudflare Access in front of it. Do not expose Flask directly to the public internet.

See:

- `cloudflare_setup_notes.md`
- `public_deploy_checklist.md`
- `SECURITY.md`
