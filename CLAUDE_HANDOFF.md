# Claude Handoff: Inquiry Voice Form

## Project Summary

This is a local/enterprise web app that turns Japanese voice, typed text, or an uploaded order-form image into a completed inquiry Excel workbook.

The app lives at:

```text
C:\Users\koki0\inquiry_voice_form
```

Main file:

```text
app.py
```

## User Goal

The user wants to complete `引合書+値引` without typing exact customer/product codes.

Desired workflow:

1. Speak or type natural Japanese.
2. App extracts:
   - 得意先
   - 納入先
   - 受注日
   - 生産日
   - 出荷希望日
   - 倉庫
   - 値引名
   - 値引方法
   - 値引率
   - 製品名/製品コード
   - 数量
3. App suggests master-data matches.
4. User confirms ambiguous candidates.
5. App generates a new Excel file.

## Current Implementation

Backend:

- Flask app in `app.py`.
- Excel editing via `openpyxl`.
- Gemini text parsing via `generateContent`.
- Gemini Vision image parsing via `generateContent` with inline image data.
- Local parser fallback for text input.
- Suggestion API for direct form typing.

Frontend:

- Single embedded HTML string in `app.py`.
- Browser Web Speech API for Japanese voice input.
- Manual text input.
- Order image upload.
- Candidate dropdowns for customer, delivery, and product.

## Current Routes

- `GET /`
- `GET /api/status`
- `POST /api/parse`
- `GET /api/suggest?kind=customer|delivery|product&q=...`
- `POST /api/parse-image`
- `POST /api/save`
- `GET /download/<token>`

## Current Environment Variables

- `GEMINI_API_KEY`: required for Gemini text/image parsing.
- `GEMINI_MODEL`: defaults to `gemini-2.5-flash`.
- `APP_PASSWORD`: required when `PUBLIC_MODE=true`.
- `INQUIRY_TEMPLATE`: optional path to real Excel template.
- `OUTPUT_DIR`: optional generated Excel output directory.
- `LOG_DIR`: optional audit log directory.
- `MAX_UPLOAD_MB`: defaults to `12`.
- `OUTPUT_RETENTION_MINUTES`: generated file retention.
- `DOWNLOAD_TOKEN_MINUTES`: download token lifetime.
- `PUBLIC_MODE`: production safety mode.
- `REQUIRE_HTTPS`: blocks non-HTTPS requests when enabled.

## Important Security Decision

Do not commit:

- `.env`
- generated Excel files
- logs

The user explicitly approved including the real workbook data for this handoff. Therefore `template.xlsx` may be committed, but the GitHub repository must be treated as private/confidential unless the user separately approves public release.

Still do not commit API keys, passwords, `.env`, generated Excel files, or logs.

## Excel Mapping

Target sheet:

```text
引合書+値引
```

Key cells currently written:

- `A4`: 得意先コード
- `A6`: 納入先コード
- `I4`: 受注日
- `H5`: 生産日 or `在庫`
- `I6`: 出荷希望日
- `B17`: 倉庫
- `L7`: 値引名
- `L8`: 値引方法
- `N8`: 値引率
- `A11`: 商品コード
- `E11`: 数量
- `I11`: order note/order number if available

## Business Rules

- 受注日 omitted -> today.
- 数量 omitted -> `1`.
- 生産日 omitted -> `在庫`.
- 出荷希望日 omitted and 生産日 provided -> production date + 2 business days, skipping Saturday/Sunday.
- 倉庫 -> `011` or `031`.
- 値引方法 -> `外掛` or `内掛`.
- Similar customer/delivery/product matches should be shown to the user for confirmation, especially when names exist in multiple prefectures.

## Known Recent Fixes

- Gemini API now reads `.env` using `utf-8-sig` because PowerShell wrote BOM.
- Default Gemini model changed to `gemini-2.5-flash`.
- `template.xlsx` was copied into the app folder for portability, but it is ignored by Git.
- UI status now calls `/api/status` instead of always saying Gemini API is required.

## Multi-PC Deployment Recommendation

Preferred:

```text
Other PCs -> browser -> Cloudflare Access/VPN/reverse proxy -> one app host
```

Avoid distributing `.env`, Gemini key, and the real Excel template to every PC.

For Cloudflare, see:

```text
cloudflare_setup_notes.md
```

For direct public HTTPS, see:

```text
public_deploy_checklist.md
```

## Next Best Tasks

1. Move HTML out of `app.py` into templates/static files.
2. Add automated tests for:
   - date rules
   - business-day shipping
   - local parser fallback
   - suggestion ranking
   - Excel cell writing
3. Replace Basic Auth with named users if this becomes team-wide.
4. Add a real deployment script for Windows Task Scheduler or Windows Service.
5. Decide whether the GitHub repository will be private and whether `template.xlsx` can be stored in Git LFS/private release assets.

## How To Start Locally

```powershell
cd C:\Users\koki0\inquiry_voice_form
powershell -ExecutionPolicy Bypass -File .\start_local.ps1
```

Then open:

```text
http://127.0.0.1:8765
```
