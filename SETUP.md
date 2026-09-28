# STARFILL Lucky Spin — Setup & Event Guide

The slot machine (`index.html`, hosted on GitHub Pages) talks to a small
backend that runs as a Google Apps Script attached to your Google Sheet
(`google-apps-script.gs`). **The backend decides every spin**, so the rules
hold on every phone and tablet, including after a reload, in incognito, or on a
second device:

| Rule | Where to change it (`google-apps-script.gs` → `CONFIG`) |
|---|---|
| 3 spins per mobile number per day | `MAX_SPINS_PER_DAY` |
| A winner can't play again for 24 hours | `WIN_LOCK_HOURS` (set 48 to block winners for the whole 2-day event) |
| Exactly 2 winners per day, never more | `WINNERS_PER_DAY` |
| Winners spread over the first ~120 spins (~42 visitors) | `PRIZE_WINDOW_SPINS` |
| Any prize still left after 15:00 IST goes to the next spin (stall open 09:30–17:00 IST, 3–4 Oct) | `LAST_CALL_TIME` |
| Prize mix when someone wins: PLUS 50%, DEEP+ 30%, IMPLANT+ 20% | `SYMBOLS[].weight` |
| A phone must wait 2 min before checking in a *different* number (booth tablet exempt) | `NEW_NUMBER_COOLDOWN_SECONDS` |

Every check-in and spin is written to the Google Sheet (tabs **Registrations**
and **Spins**). The admin page downloads a per-day Excel file.

---

## 1. Update the Apps Script (keep the same URL)

1. Open the Google Sheet → **Extensions → Apps Script**
2. Select everything in the editor, delete it, and paste the whole of
   `google-apps-script.gs` from this repo. Click **Save** 💾
3. **Set the admin password:** ⚙️ **Project Settings** → scroll to
   **Script Properties** → **Add script property**
   - Property: `ADMIN_KEY`
   - Value: a new password that nobody else knows. Don't reuse `starfill2026`,
     because it was in the public repo.
4. **Redeploy without changing the URL:** **Deploy → Manage deployments** →
   ✏️ (edit) on the existing deployment → **Version: New version** → **Deploy**.
   (Don't use "New deployment", which gives a new URL.)
   Keep **Execute as: Me** and **Who has access: Anyone**.
5. If Google asks you to authorize again, go through **Advanced → Go to (project) → Allow**.

> First time setting up? Create a sheet at [sheets.new](https://sheets.new),
> do steps 1–3, then **Deploy → New deployment → Web app** (Execute as **Me**,
> access **Anyone**), and paste the `/exec` URL into `API_URL` in `index.html`.

## 2. Test it (before the event)

1. Open the live site, check in with your own number, and spin 3 times.
2. Check in again with the same number. It must say the spins are used up.
3. Admin: open `https://<your-site>/?admin`, then enter the `ADMIN_KEY` password.
   You should see today's row, and **⬇ Excel** should download an `.xlsx`.
4. Optional test win: while signed in to the admin page, change the URL
   **in the same tab** to `https://<your-site>/?forceWin=implant` and play with
   a spare number. Test wins show as `TEST WIN` and don't use up a real prize.

## 3. Reset before the conference opens ⚠️

Testing creates real records. Before day 1, in the Apps Script editor, pick
**`resetAllData`** in the function dropdown and click **Run**. It clears
all spin counters and archives the test tabs.

## 4. During the event

- **Admin page:** `/?admin`. Shows players, spins and winners per day
  (e.g. `1 / 2`), with **⬇ Excel** for each day and **Excel — all days**.
  Each file has a *Players* sheet (name, mobile, spins used, result, prize,
  claim code) and a *Spins* sheet (every spin).
- **Handing over a prize (do this every time):**
  1. The winner shows a claim code such as `SFI-7KQ4XM9`. Type it into
     **Verify a claim code** on the admin page.
  2. Check the number is really theirs, because nobody verifies numbers at check-in:
     - the WhatsApp claim reached the office **from that same number**, or
     - give that number a missed call and watch the winner's phone ring.
  3. If it matches, tap **✔ Verified — mark prize given**. The code can't be
     used again.
  4. If it doesn't match, tap **✖ Not their number — void**. The prize goes back
     into that day's pool, so the day still ends with 2 real winners
     (usually the very next spin wins).
- **Booth tablet:** open `/?admin` **on the tablet itself**, sign in, and tap
  **Mark this device as booth tablet**. Other phones must wait 2 minutes
  before checking in a different number (to stop people cycling fake numbers).
  The booth tablet is exempt, so visitors can play one after another.
  Sign out of admin on the tablet afterwards.
- **WhatsApp:** winners tap **Claim on WhatsApp**, which sends their name,
  mobile and claim code to the office number (+91 99872 64974). Everyone else
  gets a **Chat with us on WhatsApp** button at the end.
- **Internet is required.** If the connection drops, the game says so and the
  player can tap SPIN again. A spin is never lost or counted twice.
- **Shared booth tablet:** tap **FINISH** (or **EXIT**) to return to the
  check-in screen for the next visitor.
