# Inquiry Voice Form

Japanese voice/photo input web app for generating inquiry Excel files from an existing workbook template.

## What It Does

- Accepts Japanese voice or typed text such as:
  `得意先 良栄社、受注日 今日、出荷希望日 明後日、倉庫011、値引き外掛け、値引き率3%、製品SP853A`
- Uses Gemini API when configured, with local parsing fallback for text.
- Uses Gemini Vision for order-form image recognition.
- Suggests customer, delivery destination, and product candidates from the workbook master sheets.
- Handles ambiguous names by showing candidates for user confirmation.
- Generates a new Excel file from `引合書+値引`.

## Current Business Rules

- If order date is omitted, use today.
- If quantity is omitted, use `1`.
- If production date is omitted, write `在庫` to the production-date cell.
- If production date is provided and shipping date is omitted, shipping date is two business days after production date. Saturdays and Sundays are skipped.
- Warehouse is `011` or `031`.
- Discount method is `外掛` or `内掛`.
- Kubota `注文書（出荷指示書）` forms are handled with dedicated overrides:
  customer `61110`, `大網営業所` delivery `61110005`, and model-name product matching such as `RM953X/K` -> `0343-0210`.

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
