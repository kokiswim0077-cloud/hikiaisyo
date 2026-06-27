# Public internet deployment checklist

This is the safer direct-public setup:

Internet -> Router TCP 443 -> This PC Caddy HTTPS -> Waitress on 127.0.0.1:8765 -> Flask app

Do not expose Flask or Waitress directly to the internet.

## One-time requirements

1. Buy or prepare a domain.
2. Create a DNS A record for the domain/subdomain pointing to your home global IP.
   Example: `inquiry.example.com -> your home IP`.
3. On the home router, forward only TCP 443 to this PC.
4. Keep Windows firewall closed except for Caddy on TCP 443.
5. Set app password:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\set_app_password.ps1
   ```
6. Set Gemini API key:
   ```powershell
   powershell -ExecutionPolicy Bypass -File .\set_gemini_key.ps1
   ```

## Start production

```powershell
cd C:\Users\koki0\inquiry_voice_form
powershell -ExecutionPolicy Bypass -File .\start_public_stack.ps1 -Domain inquiry.example.com
```

Login:

```text
username: user
password: APP_PASSWORD you set
```

## Security controls already implemented

- Basic authentication is required in `PUBLIC_MODE`.
- HTTPS is required in `PUBLIC_MODE` stack.
- Flask development server is not used; Waitress serves the app.
- Caddy terminates HTTPS and renews certificates.
- Upload size limit defaults to 12 MB.
- Upload file type is restricted to jpg/jpeg/png/webp/pdf.
- Rate limits are applied per client IP.
- Download links use random one-time-style tokens instead of filenames.
- Generated Excel files are deleted after `OUTPUT_RETENTION_MINUTES` (default 120).
- Download tokens expire after `DOWNLOAD_TOKEN_MINUTES` (default 60).
- Browser caching is disabled.
- Security headers are set.
- Audit logs rotate under `logs/audit.log`.

## Remaining recommendations

- Use a domain only for this app, not a broad company domain.
- Use a long random APP_PASSWORD stored in a password manager.
- Rotate the Gemini API key periodically.
- Restrict the Gemini API key in Google Cloud if possible.
- Back up `template.xlsx` separately.
- Do not leave old generated Excel files in `outputs`.
- Consider adding IP allowlisting if your office global IP is stable.
