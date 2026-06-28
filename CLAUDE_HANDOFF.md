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

## Kubota Order-Form Rule

The user provided sample images of Kubota `注文書（出荷指示書）` forms.

For this format:

- If the image is a Kubota shipping instruction form and refers to `関東甲信クボタ`, force customer code `61110`.
- If the destination office is `大網営業所` or the Kubota office code is `040`, force delivery code `61110005`.
- Delivery name should resolve to `㈱関東甲信クボタ　大網営業所`.
- Read product from `形式名` / model name such as `RM953X/K`.
- `RM953X/K` resolves to product code `0343-0210`.
- Read shipping date from the lower remarks area, e.g. `6/29(月)出荷`.
- Read warehouse from the lower remarks area, e.g. `011倉庫`.
- Warehouse is based on the shipping source text, not the customer name. `オーレック関物`, `関東物流`, `関物`, or `関東物` means `031`; `福岡倉庫` or OREC head-office/Fukuoka wording means `011`.
- If OCR or an old note reads `033`, convert it to `031`; `033` should not be shown or written to Excel.
- For this image format, order date should be today, not necessarily the document's printed order date.

Validated sample result:

- Customer: `61110`
- Delivery: `61110005`
- Product: `0343-0210`
- Warehouse: `011`
- Shipping date: `2026-06-29`

## Additional Image Rules Learned

Validated on four real sample images on 2026-06-28.

1. Kubota shipping instruction, Ichihara office:
   - Kubota office code `038` or `市原営業所` -> delivery `61110004`.
   - Model `WMC747AP/M` -> product `0358-0120`.
   - Lower remarks can include both shipping and production dates:
     `7/7(火)出荷`, `7/3(金)生産`.

2. OREC inquiry sheet / 良栄社商会:
   - Customer/delivery code `61376`.
   - Product `HR403` can resolve to `0365-0020`.
   - Handwritten note such as `24日出荷希望です` should drive shipping date.
   - Do not force warehouse `031` just because the customer is 良栄社商会. Use the shipping source/出荷口 text.

3. General order sheet / 木嶋商店:
   - Right-side seller/orderer `有限会社 木嶋商店` -> customer `61310`.
   - If no separate delivery is present, delivery also `61310`.
   - `RM983FX/K` -> product `0347-0210`.
   - Month-only notes should be preserved as text:
     `生産日 8月` -> `H5 = 8月`,
     `出荷予定日 2026年9月` -> `I6 = 2026年 9月`.

4. 受注・発注カード / エルタ:
   - Right-side company `株式会社エルタ` is the customer, even if the table has another `得意先名`.
   - `エルタ` -> customer `65137`.
   - If no direct delivery is present, delivery also `65137`.
   - `RCHR800A` -> product `0372-0180`.
   - `出荷予定日 2026年10月以降` should be preserved as text in the shipping cell.

5. General order sheet / 大竹産業 and OREC internal addressees:
   - `大竹産業` -> customer/delivery `61323`.
   - `宛先 オーレック`, `関東営業所`, or staff names such as `谷尾` are internal addressees and should not become customer/delivery.
   - If a customer is known and the parsed delivery is internal OREC, use the customer as the delivery destination.
   - `RCSP540` -> product `0374-0060`.

6. Voice input improvements:
   - Shipping-source words should infer warehouse: `オーレック関物`/`関東物流` -> `031`, `福岡倉庫`/`本社` -> `011`.
   - Discount rates like `値引き率3パーセント` should parse as `3`.

## Known Recent Fixes

- Gemini API now reads `.env` using `utf-8-sig` because PowerShell wrote BOM.
- Default Gemini model changed to `gemini-2.5-flash`.
- `template.xlsx` was copied into the app folder for portability, but it is ignored by Git.
- UI status now calls `/api/status` instead of always saying Gemini API is required.
- Gemini calls retry once to reduce transient image-reading failures.
- API keys are read from server environment only. Do not accept Gemini keys from browser payloads.

## Multi-PC Deployment Recommendation

Preferred:

```text
Other PCs -> browser -> Cloudflare Access/VPN/reverse proxy -> one app host
```

Avoid distributing `.env`, Gemini key, and the real Excel template to every PC.

Generated Excel files should be stored in Google Drive. The app now auto-detects common local Google Drive sync folders and saves under `引合書_見積書`; otherwise set `OUTPUT_DIR` or `GOOGLE_DRIVE_OUTPUT_DIR`. Use `set_google_drive_output.ps1` for manual setup.

For Cloudflare, see:

```text
cloudflare_setup_notes.md
```

For direct public HTTPS, see:

```text
public_deploy_checklist.md
```

## Quote Request / Estimate Feature

Added `/quote` for quote request handling, parallel to the inquiry form.

Scope:

- Read quote requests from voice/text or photos.
- Generate both an OREC-style estimate Excel and a quote-request copy Excel.
- Excluded by user request: response deadline, freight, payment terms, FAX status.

Fields kept:

- Customer name/code, staff name, quote date, product name/code, quantity, retail price, wholesale price, discounted price, discount name/rate, stock status, production date, ship date, note.

Learned quote rules:

- `伊藤産業機械` -> customer `61105`.
- `大竹産業` -> customer `61323`.
- Models such as `AM64B`, `RM832A`, `SP853A`, `RCSP540`, `AM65B` resolve through the product master.
- Ignore boilerplate estimate text such as production-date disclaimer, product-not-reserved notice, and standard greeting.
- Move `在庫限り` into stock status rather than leaving it in notes.
- Preserve business notes such as `入札案件です 6/1(月)早めに返信願います` and `8月以降で回答していい？`.

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
