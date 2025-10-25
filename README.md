# Email Automation (NCET Notes)

This is a simple bulk email sender with a web UI. It supports Excel/CSV uploads and optional HTML body (paste into the textarea) with template variables.

Features:
- Upload CSV/XLSX with Name and Email columns
- Paste HTML content to be used as email body (supports {{to_name}}, {{email}}, {{name}}, and any column names)
- Optional attachment file path per row (attachment column)
- Uses Gmail SMTP (or any SMTP server) via environment variables
- Production-ready tweaks for deployment (helmet, rate-limiting, file-size limits)

Quick start (local):

1. Install dependencies

```powershell
npm install
```

2. Create `.env` from `.env.example` and fill SMTP credentials

```text
SMTP_USER=your-email@gmail.com
SMTP_PASS=your-app-password
PORT=5000
```

3. Start server

```powershell
node server.js
```

4. Open http://localhost:5000

Deploying to Render
- Push this repo to GitHub and create a Render Web Service.
- Use `node server.js` as the start command (or `npm start`).
- Add environment variables `SMTP_USER` and `SMTP_PASS` in the Render dashboard.
- Set the build command to `npm install` and the start command to `npm start`.

Security notes
- Never commit your `.env` to source control. `.gitignore` already contains `.env` and `uploads/`.
- For production Gmail usage, consider OAuth2 or a transactional email provider to avoid deliverability issues.

If you want, I can add a small admin settings page to update SMTP credentials from the UI (stored temporarily) or implement OAuth2. 
