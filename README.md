# Spirelight Referral Form

A small static referral-intake form for the Spirelight voice-recording program.
Anyone with the link can submit a referral — no login required. Submissions
are appended as rows to the **Spirelight Referral Tracker** Google Sheet via
a Google Apps Script backend.

- Form: `index.html` (plain HTML/CSS/JS, no build step)
- Backend: `apps-script/Code.gs` (runs inside Google Apps Script, bound to the Sheet)

## Why this shape

A static page can't write to a database by itself — it needs somewhere to
send submissions. Apps Script was chosen over a third-party backend
(Supabase, Formspree, etc.) because it writes directly into the tracker
Sheet that's already the source of truth, needs no new accounts, and costs
nothing.

## One-time setup

### 1. Deploy the Apps Script backend

1. Open the **Spirelight Referral Tracker** Sheet:
   https://docs.google.com/spreadsheets/d/1IN1iv6X-isl2grAIG3f_LXHk1KrgUleqGXWmd3fdAdI/edit
2. **Extensions > Apps Script**.
3. Delete the starter `myFunction() {}` code, paste in the contents of
   `apps-script/Code.gs`, and save (Ctrl+S).
4. **Deploy > New deployment**.
   - Click the gear next to "Select type" and choose **Web app**.
   - Execute as: **Me**
   - Who has access: **Anyone**
   - Click **Deploy**.
5. Google will ask you to authorize the script — click through
   **Advanced > Go to (project name) (unsafe)**. This warning is Google's
   standard prompt for any script you haven't published to their store; it's
   your own script running against your own Sheet, so it's safe to allow.
6. Copy the **Web app URL** it gives you (ends in `/exec`).

### 2. Connect the form to it

1. Open `index.html`, find the line near the top of the `<script>` block:
   ```js
   const SCRIPT_URL = "REPLACE_WITH_YOUR_APPS_SCRIPT_WEB_APP_URL";
   ```
2. Replace the placeholder with the URL you copied.

### 3. Publish it

Push this folder to GitHub and turn on GitHub Pages
(**Settings > Pages > Deploy from a branch**, pick `main` and `/ (root)`).
The form will be live at `https://<your-username>.github.io/<repo-name>/`.

That URL is what gets shared in the WhatsApp group description, in place of
a Google Form link.

## Testing

Submit a test referral through the live page, then check the "Referidos" tab
on the tracker Sheet (created automatically on first submission) — a new row
should appear within a few seconds.

## Notes

- The form always shows a "¡Gracias!" success screen once the request is
  sent, even though the browser can't read Apps Script's response (Apps
  Script web apps don't support normal CORS reads, so the request is fired
  with `no-cors`). If the Apps Script deployment is broken, submissions will
  silently fail to save even though the user sees success — test end-to-end
  after any redeploy of the Apps Script.
- The sheet's "Referidos" tab only gets the raw submission columns. Add the
  same tracking columns used elsewhere in this program (Estado, Bono pagado,
  Fecha de pago) manually to the right of it — the script never touches
  those.
