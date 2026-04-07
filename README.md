# Booking Automation – Chrome Extension

A Chrome extension that combines hotel reservations from **booking.com** and a **Google Sheet** (phone reservations), then writes a monthly calendar view to a target Google Sheet.

## Calendar Format

The output sheet tab (named `YYYY-MM`) looks like this:

| Room    | 1     | 2     | … | 30    |
|---------|-------|-------|---|-------|
| Room 1  | Alice | Alice | … |       |
| Room 2  | Bob   | Bob   | … | Bob   |
| …       |       |       |   |       |
| Room 10 |       |       |   |       |

- **Rows** – up to 10 hotel rooms, auto-assigned using a greedy algorithm (earliest check-in first).
- **Columns** – one per day in the selected month.
- **Cells** – guest name if the room is occupied on that day; empty otherwise.
- Check-out day is treated as free (the outgoing guest is not shown on their departure day).

## Setup

### 1. Google Cloud OAuth credentials

Before loading the extension you must create an OAuth 2.0 client ID:

1. Go to the [Google Cloud Console](https://console.cloud.google.com/).
2. Create a project (or select an existing one).
3. Enable the **Google Sheets API**.
4. Navigate to **APIs & Services → Credentials → Create credentials → OAuth client ID**.
5. Select **Chrome Extension** as the application type.
6. Enter the extension's ID (shown in `chrome://extensions` after loading the unpacked extension).
7. Copy the generated **Client ID**.
8. Open `manifest.json` and replace `YOUR_GOOGLE_OAUTH_CLIENT_ID.apps.googleusercontent.com` with your client ID.

### 2. Load the extension

1. Open Chrome and go to `chrome://extensions`.
2. Enable **Developer mode** (toggle in the top-right).
3. Click **Load unpacked** and select this directory.

## Usage

1. Navigate to the booking.com hotel admin page:  
   `https://admin.booking.com/hotel/hoteladmin/extranet_ng/manage/home.html`
2. Click the extension icon to open the popup.
3. Fill in the fields:
   - **Phone Reservations Sheet ID** *(optional)* – the ID of a Google Sheet containing phone bookings.  
     Expected columns: `A = Guest name`, `B = Check-in date`, `C = Check-out date`, `D = Room (optional)`.
   - **Phone Sheet Range** – range to read, e.g. `Sheet1!A2:D` (default).
   - **Target Sheet ID** *(required)* – the Google Sheet where the calendar will be written.
   - **Month / Year** – the month you want to generate.
4. Click **🔑 Sign in** to authenticate with Google (only needed once per session).
5. Click **▶ Run** to generate and write the calendar.

The extension creates (or overwrites) a sheet tab named `YYYY-MM` in the target spreadsheet.
