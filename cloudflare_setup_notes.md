# Cloudflare Tunnel + Access setup

This app should stay bound to `127.0.0.1:8765`.
Cloudflare Tunnel will connect Cloudflare to the local app without opening a router port.

## Required Cloudflare settings

1. Log in with Cloudflare:
   ```powershell
   cloudflared tunnel login
   ```

2. Create a named tunnel:
   ```powershell
   cloudflared tunnel create inquiry-voice-form
   ```

3. Route a hostname to the tunnel.
   Replace `inquiry.example.com` with your Cloudflare-managed hostname:
   ```powershell
   cloudflared tunnel route dns inquiry-voice-form inquiry.example.com
   ```

4. Create a config file at:
   `%USERPROFILE%\.cloudflared\config.yml`

   Example:
   ```yaml
   tunnel: inquiry-voice-form
   credentials-file: C:\Users\koki0\.cloudflared\<TUNNEL_ID>.json

   ingress:
     - hostname: inquiry.example.com
       service: http://127.0.0.1:8765
     - service: http_status:404
   ```

5. In Cloudflare Zero Trust:
   - Access > Applications > Add application > Self-hosted
   - Application domain: `inquiry.example.com`
   - Policy: Allow only your email address
   - Session duration: choose a short period such as 12 hours

6. Run the tunnel:
   ```powershell
   cloudflared tunnel run inquiry-voice-form
   ```

## Security defaults used by the app

- Gemini API key stays only in `.env` on this PC.
- App has optional Basic Auth via `APP_PASSWORD`.
- Upload size is limited by `MAX_UPLOAD_MB` (default 12 MB).
- Browser caching is disabled.
- The Flask app itself is not bound to the public network.
