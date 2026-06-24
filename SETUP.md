# STARFILL Lucky Spin — Central Results Setup

Follow these steps once to make every spin (from every customer's phone) land in
one Google Sheet you can open anytime.

## 1. Create the Google Sheet

1. Go to [sheets.new](https://sheets.new) (signed in with the clinic's Google account)
2. Rename it something like **STARFILL Spins**

## 2. Add the script

1. In the sheet, click **Extensions → Apps Script**
2. Delete any starter code in the editor
3. Open `google-apps-script.gs` (in this repo), copy everything, paste it in
4. Click the **Save** icon (💾)

## 3. Deploy it as a web app

1. Click **Deploy → New deployment**
2. Click the gear ⚙️ next to "Select type" → choose **Web app**
3. Set:
   - **Description:** STARFILL logger
   - **Execute as:** Me
   - **Who has access:** **Anyone**  ← important
4. Click **Deploy**
5. Click **Authorize access** → pick your Google account → if it warns
   "Google hasn't verified this app", click **Advanced → Go to (project) → Allow**
   (it's your own script, this is safe)
6. Copy the **Web app URL** — it ends in `/exec`

## 4. Connect it to the slot machine

1. Open `index.html` in this repo
2. Near the top of the `<script>`, find the `CFG` block:
   ```javascript
   WEBHOOK_URL:  '',
   ```
3. Paste your `/exec` URL between the quotes:
   ```javascript
   WEBHOOK_URL:  'https://script.google.com/macros/s/AKfy.../exec',
   ```
4. Save, then push:
   ```bash
   git add index.html && git commit -m "chore: connect Google Sheets webhook" && git push origin main
   ```

## 5. Test

1. Open the live site, do one spin
2. Check the Google Sheet — a new row should appear within a few seconds

That's it. Every spin from any device now logs to this one sheet, with
timestamp, name, phone, the three reels, win/miss, and the prize won.

> The in-browser admin panel (`?admin=starfill2026`) still works, but it only
> shows spins from the device you open it on. The Google Sheet is the real,
> complete record — use that one.
