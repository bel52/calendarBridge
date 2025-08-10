# Mac Outlook ➜ Google Calendar (One‑Way) — v2.0

Local, script‑only sync from **Microsoft Outlook for macOS (client)** to **Google Calendar**.
**One‑way**: Outlook → Google. Runs hourly via `launchd`.

---

## Features

* **Export window:** defaults to **60 days back / 120 days ahead**
* **AppleScript exporter** → writes `.ics` files to `outbox/`
* **ICS cleaner** → strips Outlook‑specific `X-` headers
* **Google sync (Python):**

  * Inserts/updates/deletes to keep Google in step with Outlook
  * Prevents duplicates using a private tag (`extendedProperties.private.compositeUID`)
  * Computes a content hash to skip unchanged events (idempotent)
  * Detects midnight‑to‑midnight spans and posts **true all‑day** events
  * Adds a **timezone** to floating times so Google accepts them
  * **Sanitizes RRULEs** that Google otherwise rejects
* **Hourly automation** with logs in `logs/`

> **Security:** `credentials.json`, `token.json`, `logs/`, and `outbox/` are **not** in git (see `.gitignore`). Everything runs locally on your Mac.

---

## Repository layout

```
calendarBridge/
├── exportEvents.scpt                       # Outlook → outbox/*.ics (AppleScript)
├── clean_ics_files.py                      # removes X‑ headers from ICS
├── safe_sync.py                            # one‑way sync to Google
├── full_sync.sh                            # export → clean → sync + logging
├── launchd/
│   └── com.calendarBridge.full_sync.plist  # hourly at Minute=0
├── requirements.txt
├── .gitignore
└── README.md
```

> Helper (not tracked): `list_outlook_calendars.scpt` prints your Outlook calendars with index.

---

## Requirements

* macOS (Ventura or similar)
* Microsoft Outlook (desktop app)
* Python **3.9** (virtualenv at `~/calendarBridge/.venv`)
* Google Calendar API OAuth: `credentials.json` (first run generates `token.json`)

### Python packages

`requirements.txt` pins:

* `google-api-python-client`
* `google-auth-oauthlib`
* `icalendar`
* `urllib3<2` (avoids LibreSSL warnings on macOS Python 3.9)

---

## Installation

```bash
cd ~/calendarBridge
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt

# Place Google OAuth files (NOT in git):
#  - credentials.json
#  - token.json (created on first sync)
```

---

## Select your Outlook calendar

List calendars and note the index, then set `targetCalIndex` in `exportEvents.scpt`.

**If you have the helper script:**

```bash
osascript ~/calendarBridge/list_outlook_calendars.scpt
```

**One‑liner alternative:**

```bash
osascript -e 'tell app "Microsoft Outlook" to set o to {} & repeat with i from 1 to (count of calendars) by 1 \n set end of o to (i as text) & " | " & (name of calendar i as text) \n end repeat \n return o as text'
```

Edit `exportEvents.scpt` and set:

* `targetCalIndex` → your chosen calendar index
* `exportDaysBack = 60`, `exportDaysAhead = 120` (or adjust)

---

## First manual run

```bash
cd ~/calendarBridge
source .venv/bin/activate

osascript exportEvents.scpt        # Step 1: export Outlook → outbox/*.ics
python clean_ics_files.py          # Step 2: strip X‑ headers
python safe_sync.py                # Step 3: sync to Google (opens OAuth on first run)
```

* Re‑running `safe_sync.py` should show mostly **skipped** (dedupe & hashing).
* Check your Google Calendar for all‑day and recurring series correctness.

---

## Hourly automation (launchd)

```bash
mkdir -p ~/Library/LaunchAgents ~/calendarBridge/logs
cp launchd/com.calendarBridge.full_sync.plist ~/Library/LaunchAgents/

# Reload the job cleanly
launchctl bootout gui/$(id -u) ~/Library/LaunchAgents/com.calendarBridge.full_sync.plist 2>/dev/null || true
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.calendarBridge.full_sync.plist
launchctl kickstart -k gui/$(id -u)/com.calendarBridge.full_sync
```

* Runs **every hour on the hour** (`StartCalendarInterval` Minute = `0`).
* Logs: `~/calendarBridge/logs/launchd.out` and `launchd.err`.

**Verify next run**

```bash
tail -n 200 ~/calendarBridge/logs/launchd.out
launchctl list | grep com.calendarBridge.full_sync || true
```

---

## Configuration

* **Timezone for timed events:** set env var `CALBRIDGE_TZ` (default: `America/New_York`).

  * For the LaunchAgent session:

    ```bash
    launchctl setenv CALBRIDGE_TZ America/Chicago
    launchctl kickstart -k gui/$(id -u)/com.calendarBridge.full_sync
    ```
* **Google Calendar target:** edit `CAL_ID` in `safe_sync.py` (default `primary`).
* **Export window:** edit `exportDaysBack` / `exportDaysAhead` in `exportEvents.scpt`.

---

## How it avoids duplicates

* Each Google event gets a private tag: `extendedProperties.private.compositeUID` derived from the Outlook `UID` (and `RECURRENCE-ID` for exceptions).
* The sync computes a JSON hash of event content and stores it in the description; unchanged events are skipped, deleted ones are removed.

---

## Troubleshooting

* **AppleScript exporter error or wrong calendar** → verify `targetCalIndex` using the listing step above.
* **Google 400 “Missing time zone”** → handled by `safe_sync.py`; set `CALBRIDGE_TZ` if you prefer a different zone.
* **Google 400 “Invalid recurrence rule”** → Outlook RRULEs are sanitized before insert/update.
* **SSL/urllib3 warnings** → `urllib3<2` pinned in `requirements.txt`.

---

## Contributing / Branching

* Work on a feature branch, then PR into `main`.
* Version tag for this release: **`v2.0`**.

---

## License

Private/internal use until a license is added.
