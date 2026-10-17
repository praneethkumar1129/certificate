# Certificate Studio — Student Participation Certificate Generator

A local FastAPI website that reads your Excel registration sheet, creates one personalized PDF per student using the supplied participation certificate PDF, emails each certificate through Gmail SMTP, and provides a ZIP download and CSV delivery report.

## What it supports

- `.xlsx`, `.xlsm`, `.xls`, and `.csv` registration files
- The supplied schema: `Member 1 - Full Name`, `Member 1 - Email Address`, through Member 4 (the parser also checks for members up to 20)
- Uses `College / Institution Name` on the certificate
- Skips blank member slots and records missing/invalid details as skipped
- PDF preview before sending
- Explicit checkbox confirmation before sending
- ZIP of generated certificates and a CSV delivery report
- Continues processing if one email fails

## 1. Install Python on Windows

Install Python 3.11 or 3.12 from https://www.python.org/downloads/windows/ and enable **Add Python to PATH** during installation.

Open PowerShell in this project folder. You should see `app.py`, `requirements.txt`, and the `templates` folder when you run:

```powershell
dir
```

## 2. Create a virtual environment

```powershell
py -3.12 -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt
```

If PowerShell blocks activation, run this once in that terminal, then activate again:

```powershell
Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
.\.venv\Scripts\Activate.ps1
```

If `py -3.12` is unavailable, use `py -3.11` or `python` instead.

## 3. Configure Gmail securely

1. Use the Google account that should send certificates.
2. Turn on 2-Step Verification for that account.
3. Open Google's App Passwords page: https://myaccount.google.com/apppasswords
4. Create an app password for this app. Google may not offer App Passwords for every account type or account configuration.
5. Copy `.env.example` to a new file named `.env` in this project folder.
6. Put your Gmail address in `SMTP_USER` and the generated App Password in `SMTP_APP_PASSWORD`. Remove spaces from the App Password.
7. Never paste the App Password into frontend code, GitHub, screenshots, or chat. Do not use your normal Gmail password.

Example `.env`:

```env
SMTP_HOST=smtp.gmail.com
SMTP_PORT=587
SMTP_USER=youraddress@gmail.com
SMTP_APP_PASSWORD=abcdefghijklmnop
SENDER_NAME=AURA 2K26 Certificate Team
```

## 4. Start the website

From the project folder, with the virtual environment active:

```powershell
uvicorn app:app --reload
```

Open http://127.0.0.1:8000 in your browser.

Keep the PowerShell window open while using the site. Press `Ctrl+C` to stop the server.

## 5. Generate and send certificates

1. Enter the event name exactly as it should appear in the certificate.
2. Upload the Excel registration file.
3. Leave the template upload empty to use `templates/participation_template.pdf`, or upload a replacement PDF.
4. Click **Preview certificate** and check the result.
5. Click **Review & send**.
6. Check the event name and confirm you authorize sending.
7. Click **Send certificates now**.
8. Review the Sent / Failed / Skipped counts and individual rows.
9. Download the ZIP and CSV report.

**Important:** A successful SMTP response means Gmail accepted the message for sending; it does not prove delivery to the inbox. Check failures and spam folders if needed.

## Expected Excel column names

At minimum, provide member names and email addresses using columns like:

- `Member 1 - Full Name`
- `Member 1 - Email Address`
- `Member 2 - Full Name`
- `Member 2 - Email Address`
- `Member 3 - Full Name`
- `Member 3 - Email Address`
- `Member 4 - Full Name`
- `Member 4 - Email Address`
- `College / Institution Name`

Header matching ignores capitalization and repeated spaces. Each populated member row is treated as an individual recipient. The top-level `Email Address` is used only as a fallback for Member 1 if their individual email is blank.

## Certificate layout note

This starter app keeps the original PDF as the background and covers/replaces the three lines containing the name, institution, and event fields. It preserves the logo, title, fixed symposium text, and signature labels. Since the source template uses embedded fonts and exact typography, check the preview before bulk sending. If your template has a different page size or layout, the overlay coordinates in `create_certificate()` may need adjustment.

## Security and limitations

- Intended for local use on a trusted computer; do not expose this development server to the public internet.
- Student information and generated PDFs are stored temporarily in the `jobs/` directory. Delete that folder after saving the ZIP/report.
- The current starter keeps job metadata in memory, so download/report links stop working after the server restarts. Files under `jobs/` remain until you delete them.
- Duplicate records with the same name and email are deduplicated within one upload.
- This starter is synchronous; very large batches may take a while. Test with a small spreadsheet first.
- The website does not send anything during preview or review; actual email starts only after the confirmation checkbox is selected.


## Troubleshooting: `[Errno 11001] getaddrinfo failed`

On Windows this generally means DNS could not resolve the SMTP hostname. It is not necessarily an issue with the student's email address.

1. Check `.env` has these exact values (no spaces around `=`):
   ```env
   SMTP_HOST=smtp.gmail.com
   SMTP_PORT=587
   ```
2. In PowerShell run `Resolve-DnsName smtp.gmail.com`. If this fails, check internet access, VPN/proxy settings, and DNS. You can try `ipconfig /flushdns` and then repeat the DNS test.
3. If DNS resolves, run `Test-NetConnection smtp.gmail.com -Port 587`. `TcpTestSucceeded` should be `True`; otherwise the network/firewall may block SMTP.
4. After changing `.env`, restart Uvicorn with Ctrl+C and `uvicorn app:app --reload`.
5. Never put your Gmail App Password in frontend files or share it.

## Template-preservation fix

The certificate generator now writes text into the original blank fields instead of covering the center with a large white rectangle. No white overlay rectangles are painted over the page, so the watermark and background artwork remain visible. The original underscore strokes may still be visible behind or beside inserted text; this is intentional to avoid damaging the template artwork. Check the preview before bulk emailing, especially for unusually long names or event names.
