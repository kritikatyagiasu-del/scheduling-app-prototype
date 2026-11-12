# Install (about 10 minutes)

Runs free on a Google account: Google Sheet + Apps Script + Google Calendar + Gmail + Google Meet.
You can install it on your own account to test, then copy it to the practitioner's account.

## 1. Create the Sheet
1. Open **sheets.new** and name the sheet **Bookings**.
2. Go to **Extensions ▸ Apps Script**.

## 2. Paste the code
1. In the editor, open **Code.gs**, select everything, and replace it with the contents of the `dist/Code.gs`.
2. Click the **gear (Project Settings)**, tick **"Show appsscript.json manifest file in editor"**, then go back to the **Editor** (`< >` icon).
3. Open **appsscript.json** and replace its contents with the `dist/appsscript.json`. This turns on Google Calendar and Meet and sets India time.
4. Press **Ctrl/Cmd + S** to save.

## 3. Run setup once
1. In the function dropdown at the top, choose **setup**, then click **Run**.
2. Google asks for permissions. Click **Review permissions**, choose your account, then **Advanced ▸ Go to (project name) ▸ Allow**.
   The "unverified app" warning appears because this is your own private script. The app asks to manage your sheets and calendars, send email as you, and run on a timer for reminders.
3. Setup creates the **Settings / Patients / Bookings / Log** tabs, a separate **"<practice name> – Bookings"** calendar, an admin PIN, and a reminder timer that runs every 5 minutes.

## 4. Publish the web app
1. Click **Deploy ▸ New deployment**, then the gear ▸ **Web app**.
2. Set **Execute as: Me** and **Who has access: Anyone**, then click **Deploy**.
3. Copy the **Web app URL** (it ends in `/exec`).

## 5. Fill in Settings (in the Sheet)
| Setting | What to put |
|---|---|
| appUrl | The `/exec` URL from step 4 |
| therapistEmail | Where the therapist's alerts go |
| therapistPhone | e.g. +91 98xxxxxxxx (shown to patients for same-day changes) |
| clinicAddress | For in-person sessions |
| upiId / upiName | UPI ID for payments |
| workDays, dayStart, dayEnd | Working hours |
| bufferMinutes, maxPerDay | Room for long sessions (default: 2 h kept free on each side, max 2 a day) |
| privacyText | Replace the draft with lawyer-reviewed text |

Reload the Sheet, then go to **Booking app ▸ Show admin PIN and links**.

## 6. Use it on the phone
- **Patients:** the `/exec` link.
- **Therapist:** the `/exec?page=admin` link. Open it on their phone, enter the PIN, then use the browser menu ▸ **Add to Home screen**.
- They can also block time by adding events to the bookings calendar in the Google Calendar app.

## Updating the code later
1. Paste the new `Code.gs` over the old one and save.
2. If the update notes say so, choose **setup** and click **Run** again. It's safe to re-run: it keeps your settings, PIN, calendar and data.
3. Go to **Deploy ▸ Manage deployments ▸ pencil ▸ Version: New version ▸ Deploy**. The link stays the same.

## Dashboard
- **In the app:** the **Insights** tab has a month picker and shows total, new and returning patients, earnings (UPI/cash), sessions, no-shows, cancellations and payments due.
- **In the Sheet:** the **Dashboard** tab shows the same numbers. Pick a month in cell B3; there's a month-by-month table below. It updates automatically when anything changes, or use **Booking app ▸ Refresh dashboard**.

## Testing tips
- Make a test booking from your phone with a second email address.
- To try reminders and "Session done" without waiting a day, set **minNoticeHours** to 0 temporarily, book something 2½ hours ahead, then set it back.
- Emails go out from the Gmail account that installed it. Only test with people who know.

## Moving it to another Google account later
1. In the new Google account: open this Sheet (shared with that account) ▸ **File ▸ Make a copy**. The script is copied with it.
2. In the copy, clear the test rows in **Patients**, **Bookings** and **Log**. Clear the **calendarId** and **adminPin** values so setup creates fresh ones.
3. Set **busyCalendarIds** to `primary` so the main calendar also blocks slots.
4. Repeat steps 3 to 6 in that account. Emails then go out from that Gmail, and Meet links belong to that account.

## Known limits
- **Banner:** Google shows a small "created by a Google Apps Script user" banner at the top of the pages. It's normal for free accounts.
- **Email cap:** free Gmail sends to 100 recipients a day through scripts. That's about 14 sessions a day.
- **Meet:** patients join the Meet link and the therapist admits them. Free accounts allow one-to-one calls up to 24 hours.
