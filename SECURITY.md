# Security Notes

## Do Not Commit

- `.env`
- Gemini API keys
- `APP_PASSWORD`
- Real customer/product Excel template files, unless the repository is private and the user explicitly approves including them
- Generated Excel output files
- Logs containing operational metadata

## Current Handoff Decision

The user approved including real workbook data for Claude/GitHub handoff. Treat any repository containing `template.xlsx` as private/confidential.

Still do not commit `.env`, Gemini API keys, passwords, generated Excel files, or logs.

## Recommended Deployment

Best option for multiple PCs:

```text
Users' browsers -> Cloudflare Access or VPN -> one internal app host -> 127.0.0.1:8765 Flask/Waitress
```

Do not install the Gemini key and real Excel template on every PC unless there is a strong reason.

## Existing Controls

- Optional Basic Auth through `APP_PASSWORD`.
- `PUBLIC_MODE=true` refuses to run without `APP_PASSWORD`.
- Optional HTTPS enforcement through `REQUIRE_HTTPS`.
- Upload size limit through `MAX_UPLOAD_MB`.
- Voice/text input length limit through `MAX_TEXT_CHARS`.
- Gemini API keys are read from server environment only, not from browser/API payloads.
- Upload extension/MIME validation for images/PDFs.
- Upload file signature validation for jpg/png/webp/pdf.
- Generated downloads use random tokens instead of raw filenames.
- Generated Excel retention cleanup.
- Browser cache disabled.
- Audit log support.

## Remaining Hardening Work

- Add per-user login instead of shared Basic Auth if more than a few users will use it.
- Add CSRF protection if cookie/session authentication is introduced.
- Add IP allowlisting or Cloudflare Access policies.
- Encrypt stored templates and outputs if the machine is shared.
- Add a scheduled cleanup task for output/log retention.
- Add monitored backups for the Excel template only.

## Google Drive Storage

Generated Excel files may be stored under a local Google Drive sync folder, usually `引合書_見積書`.

- Keep this folder in My Drive or a restricted shared drive.
- Do not make the folder public or link-shared broadly.
- Use Google Drive permissions for access control, not file-name secrecy.
- If multiple PCs use the app, prefer one app host writing to one Google Drive location.
