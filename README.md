# CalendarBridge — v4.0.0 (Stable)

One‑way, local **Outlook (macOS client) → Google Calendar** sync that you fully control.  
No Exchange Online admin visibility; runs locally via AppleScript + Python + Google Calendar API.

## What’s new in v4
- Stable iCalUIDs (uses Outlook UID when available + UTC instance start) → prevents duplicate/cancelled churn.
- Accurate mirroring: add / update (summary, location, description) / delete orphans.
- Optional Google **batch mode** for faster create/update/delete (reduced quota pressure).
- Safer ICS parsing: strips problematic `X-` headers (e.g., `X-ENTOURAGE_UUID`) before parsing.
- Robust exponential backoff & rotating logs.

> Design: one‑way sync from Outlook to Google. If you move a meeting time in Outlook, a new UID is generated for that instance; the old Google instance is deleted and the new one is created.

---

## Folder layout (standardized)
$HOME/calendarBridge/
├── .venv/ # Python venv (local only)
├── outbox/ # AppleScript export target (.ics)
├── logs/ # Logs
├── exportEvents.scpt # AppleScript: export Outlook → ICS
├── calendar_sync.py # Main sync logic (v4)
├── sync.sh # Entry point used by launchd + manual runs
├── calendar_config.json # Config (IDs, windows, batch on/off, TZ)
├── VERSION # 4.0.0
└── (credentials.json, token.json) # Google OAuth (kept local; gitignored)

---

## Requirements
- macOS (Ventura or similar), AppleScript enabled
- Python 3.9 (recommended venv at `$HOME/calendarBridge/.venv`)
- Google API OAuth credentials saved locally as `credentials.json` (do **not** commit)
- Python packages (pin in your requirements file):
  - `google-api-python-client`, `google-auth`, `google-auth-oauthlib`
  - `icalendar`, `recurring_ical_events`, `pytz`

Create/activate venv & install:
```bash
cd "$HOME/calendarBridge"
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
Configuration (calendar_config.json)
Replace placeholders with your values; keep everything else as shown.
{
  "google_calendar_id": "<YOUR_GOOGLE_CALENDAR_ID>",
  "outlook_calendar_name": "<YOUR_OUTLOOK_CALENDAR_NAME>",
  "outlook_calendar_index": 1,
  "timezone": "America/New_York",
  "sync_days_past": 90,
  "sync_days_future": 120,

  "enable_batch_operations": true,
  "batch_size": 50,
  "api_delay_seconds": 0.05
}
Notes
Updates are detected via summary, location, description.
Time changes are handled via UID regeneration → old instance deleted, new created.
Running manually
cd "$HOME/calendarBridge"
./sync.sh
Watch logs:
tail -n 200 "$HOME/calendarBridge/logs/calendar_sync.log"
Hourly automation (launchd)
Recommended LaunchAgent (at $HOME/Library/LaunchAgents/net.leathermans.calendarbridge.plist) runs at the top of every hour and at login:
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN"
 "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key><string>net.leathermans.calendarbridge</string>

  <key>ProgramArguments</key>
  <array><string>$HOME/calendarBridge/sync.sh</string></array>

  <key>WorkingDirectory</key><string>$HOME/calendarBridge</string>

  <!-- Run at the top of every hour -->
  <key>StartCalendarInterval</key>
  <dict><key>Minute</key><integer>0</integer></dict>

  <key>RunAtLoad</key><true/>

  <key>StandardOutPath</key><string>$HOME/calendarBridge/logs/launchd.out</string>
  <key>StandardErrorPath</key><string>$HOME/calendarBridge/logs/launchd.err</string>
</dict>
</plist>
Load / verify:
launchctl bootout gui/$(id -u)/net.leathermans.calendarbridge 2>/dev/null || true
launchctl load "$HOME/Library/LaunchAgents/net.leathermans.calendarbridge.plist"
launchctl list | grep net.leathermans.calendarbridge
Troubleshooting
No events created: check logs/calendar_sync.log; confirm credentials.json/token.json exist locally.
Frequent “Updating:” lines with no changes: ensure you’re on v4 (it no longer compares start/end for updates).
Cancelled duplicates: typically unstable UIDs; v4 fixes this. Consider a one‑time cleanup of old cancelled artifacts.
LibreSSL warning: harmless with macOS system Python; use venv + pinned deps.
Security / privacy
OAuth creds (credentials.json, token.json) are local only and gitignored.
No server component; nothing visible to Exchange admins or any third party.
Versioning
VERSION file contains 4.0.0.
Tag releases in Git (e.g., v4.0.0).
