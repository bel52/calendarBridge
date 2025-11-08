Updated CalendarBridge Guide
Overview
CalendarBridge is a local, one‑way synchronisation tool that exports events from an Outlook calendar on macOS and mirrors them into Google Calendar. There is no server component—everything runs on your Mac—and your Outlook and Google credentials stay on your machine.
This updated guide reflects improvements made after version 4.0.0. The previous calendar_sync.py has been replaced with a more robust safe_sync.py that generates a stable event identifier for every occurrence in your Outlook calendar. The identifier is a composite of the Outlook event UID and its start time, which prevents duplicates and ensures that recurring events are managed correctly. A new full_sync.sh script orchestrates the export, cleaning, and sync steps, and a cleanup_old_events.py script can remove legacy duplicates created by older versions.
Key improvements
Stable composite IDs – event IDs are derived from the Outlook UID and the instance start, so every occurrence has a unique identifier. This eliminates duplicate or stale events in Google.
Accurate mirroring – creates events that appear in Outlook, updates events whose details change, and deletes events that have been removed from Outlook.
Timezone‑aware sync window – the sync window is anchored in your configured timezone so that all occurrences within the past/future window are included, even if a recurring series began outside that range.
Rate‑limit handling – exponential backoff and configurable delay (api_delay_seconds) minimise API quota issues.
Batch operations (optional) – you can enable batch mode in calendar_config.json to group operations and reduce API calls.
Cleanup script – if you previously used an older version of CalendarBridge and have duplicates in Google, the cleanup_old_events.py script can remove them in bulk.
Folder layout
Your CalendarBridge folder is organised as follows (replace <HOME> with your actual home directory, e.g. /Users/bel):
<HOME>/calendarBridge/
├── .venv/                      # Python virtual environment (local only)
├── outbox/                     # AppleScript export target (.ics files)
├── logs/                       # Logs and exported run output
├── exportEvents.scpt           # AppleScript: export Outlook → .ics
├── clean_ics_files.py          # Strips problematic X‑ headers & splits multi‑event .ics
├── safe_sync.py                # Main sync logic (composite ID version)
├── full_sync.sh                # Entrypoint used by launchd & manual runs
├── cleanup_old_events.py       # Optional script to delete legacy duplicates
├── calendar_config.json        # Configuration (IDs, window lengths, batch options, TZ)
├── requirements.txt            # Python package requirements
├── VERSION                     # Version string
└── (credentials.json, token.json)   # Google OAuth files (local only; do not commit)
Requirements
macOS (Ventura or later recommended) with AppleScript enabled.
Python 3.11 or newer. Python 3.9 is end‑of‑life and triggers warnings from Google’s libraries. Install the latest Python from https://www.python.org/downloads/ and create a virtual environment.
Google OAuth credentials stored locally as credentials.json. These are never committed to Git. The first time you run the sync, a token.json will be created after you authorise access in your browser.
Outlook for macOS. CalendarBridge exports from the on‑device Outlook client via AppleScript.
Setting up the Python environment
Open Terminal and navigate to your CalendarBridge folder:
cd "<HOME>/calendarBridge"
Create a virtual environment using your installed Python (replace python3.11 with the actual path if necessary):
python3.11 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
Copy your downloaded OAuth client secret to credentials.json in this folder. On the first run, you will be prompted to grant CalendarBridge access to your Google calendar; this will generate token.json automatically.
Installing dependencies
The requirements.txt file lists all necessary packages, including:
google-api-python-client, google-auth, google-auth-oauthlib, google-auth-httplib2
icalendar, recurring-ical-events, pytz, python-dateutil
click, tzdata and other supporting libraries
You can update dependencies by editing requirements.txt and re‑running pip install -r requirements.txt.
Configuration – calendar_config.json
The calendar_config.json file controls which calendars to sync, the time window, and API tuning. Example:
{
  "google_calendar_id": "brett@leathermans.net",           // target calendar in Google
  "outlook_calendar_name": "Calendar",                    // name of the calendar in Outlook
  "outlook_calendar_index": 2,                              // index if multiple calendars share the name
  "timezone": "America/New_York",                         // your local timezone
  "sync_days_past": 60,                                     // include events up to X days in the past
  "sync_days_future": 90,                                   // include events up to Y days in the future

  "enable_batch_operations": false,                         // set true to use Google batch API
  "batch_size": 50,                                         // number of operations per batch
  "api_delay_seconds": 0.05                                 // delay between API calls when not batching
}
Notes:
If a recurring series started outside the past/future window, all occurrences inside the window will still be synced.
Changing the window will add or delete events accordingly on the next run. Larger windows mean more events and potentially more API calls.
Batch operations can significantly reduce rate‑limit errors; however, not all Google accounts allow batching. Start with enable_batch_operations: false and increase api_delay_seconds if you see many rateLimitExceeded messages.
Manual run
To run a full sync manually (export, clean, sync), use:
cd "<HOME>/calendarBridge"
source .venv/bin/activate
./full_sync.sh
The script performs these steps:
Uses exportEvents.scpt to export your Outlook calendar to a single .ics file in outbox/.
Runs clean_ics_files.py to split the exported file into individual cleaned .ics files and remove problematic headers.
Calls safe_sync.py, which:
Parses the cleaned .ics files for all occurrences inside your configured window.
Generates a composite key (UID|start) for each occurrence and a deterministic Google event ID.
Adds new events, updates existing events, and deletes events that are no longer present.
Stores mapping information in sync_state.json so that subsequent runs only affect managed events.
Logs are written to logs/full_sync_<timestamp>.log for each run. Review these logs if you encounter errors.
Dealing with legacy duplicates
If you used an earlier version of CalendarBridge, you may have duplicate events with outdated IDs. After installing the new scripts, a one‑time cleanup may be necessary:
The original sync_state.json is renamed to sync_state_backup.json during migration. The new safe_sync.py uses a composite key and starts fresh.
To delete the old events from Google Calendar, run cleanup_old_events.py:
cd "<HOME>/calendarBridge"
source .venv/bin/activate
python cleanup_old_events.py
The script reads the backed‑up state and deletes the old events, retrying on rate limits. Review the output carefully; deletions are permanent.
Hourly automation (launchd)
To run the sync automatically every hour—even after your Mac reboots—create a LaunchAgent. This agent will load at login and run the sync at regular intervals.
Create the LaunchAgent plist:
cat > "$HOME/Library/LaunchAgents/net.leathermans.calendarbridge.plist" <<'PLIST'
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
  <dict>
    <key>Label</key>
    <string>net.leathermans.calendarbridge</string>

    <!-- Use bash to ensure the virtual environment activates correctly -->
    <key>ProgramArguments</key>
    <array>
      <string>/bin/bash</string>
      <string>/Users/YOURUSER/calendarBridge/full_sync.sh</string>
    </array>

    <!-- Run every 60 minutes -->
    <key>StartInterval</key>
    <integer>3600</integer>

    <!-- Run immediately at login -->
    <key>RunAtLoad</key>
    <true/>

    <key>StandardOutPath</key>
    <string>/Users/YOURUSER/calendarBridge/logs/launchd.out</string>
    <key>StandardErrorPath</key>
    <string>/Users/YOURUSER/calendarBridge/logs/launchd.err</string>
  </dict>
</plist>
PLIST
Replace YOURUSER with your macOS username and adjust the paths if your calendarBridge folder lives elsewhere.
Load the agent:
launchctl bootout gui/$(id -u)/net.leathermans.calendarbridge 2>/dev/null || true
launchctl load "$HOME/Library/LaunchAgents/net.leathermans.calendarbridge.plist"
launchctl list | grep calendarbridge
The StartInterval key ensures the sync runs roughly every hour when your Mac is awake. RunAtLoad makes it run immediately at login. If your Mac sleeps through an interval, the next run will occur after it wakes.
To modify or unload the agent later, edit the plist and run launchctl bootout gui/$(id -u)/net.leathermans.calendarbridge to remove it.
Updating and contributing
This repository uses Git for version control. To commit your changes and push them to GitHub:
cd "<HOME>/calendarBridge"
git init                          # if the .git folder was removed
git remote add origin https://github.com/bel52/calendarBridge.git   # add your remote
git checkout -b main             # or checkout your default branch

# Add your updated files (safe_sync.py, full_sync.sh, requirements.txt, cleanup_old_events.py, README.md)
git add safe_sync.py full_sync.sh cleanup_old_events.py requirements.txt README.md calendar_config.json
git commit -m "Refactor sync with composite IDs, update scripts and documentation"

# Pull remote changes first if the repo already exists
git pull origin main --rebase

git push -u origin main
You may need to authenticate with a personal access token when pushing. If you use a different branch or wish to open a pull request, adjust the commands accordingly.
Troubleshooting
Missing events: Confirm that the event appears in outbox/clean_*.ics after exporting. If it does, verify that your sync window in calendar_config.json covers the event’s date. The script parses all occurrences within the window. Recurring events that start outside the window will still sync if an occurrence falls inside it.
Duplicate/triplicate events: Typically caused by migrating from an older version. Run cleanup_old_events.py once, then allow the new sync to recreate the proper events. Do not revert to the old script.
Rate limit errors: The sync uses exponential backoff, but you may still see rateLimitExceeded warnings. Increase api_delay_seconds or enable batch operations with a moderate batch size (e.g. 50) in your config.
LaunchAgent doesn’t run: Confirm the absolute paths in the plist are correct and that the script is executable (chmod +x full_sync.sh). Use launchctl list | grep calendarbridge to see its status. Check logs/launchd.err for any errors.
Authentication failures: Delete token.json and re‑run the sync. A browser window will open asking you to re‑authorise.
Security and privacy
Your credentials.json (Google OAuth client secret) and token.json (access token) remain on your Mac. They are not committed to version control. CalendarBridge does not send any data to external servers except to Google’s calendar API during synchronisation.
Versioning
The VERSION file contains a single line with the current version. Bump this value when you make a release, tag the commit (e.g. git tag v5.0.0), and push the tag to GitHub. This README reflects the improvements after v4 and can be considered v5.
