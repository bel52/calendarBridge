# CalendarBridge v7.0.0 — New Mac Deployment

## Prerequisites

- macOS with Microsoft Outlook in **Legacy mode** (Revert to Legacy Outlook)
- Python 3.11+ installed
- Google OAuth `credentials.json` (copy from old Mac or generate fresh)

## Step-by-Step Deployment

### 1. Clone and setup

```bash
cd ~
git clone https://github.com/bel52/calendarBridge.git
cd calendarBridge
```

### 2. Create Python virtual environment

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

### 3. Add Google credentials

Copy `credentials.json` from your old Mac:

```bash
# From old Mac, copy ~/calendarBridge/credentials.json to new Mac ~/calendarBridge/
```

Or generate fresh from Google Cloud Console → APIs & Services → Credentials.

### 4. Create outbox directory

```bash
mkdir -p ~/calendarBridge/outbox
mkdir -p ~/calendarBridge/logs
```

### 5. Verify Outlook is in Legacy mode

Open Outlook. If it's in New Outlook, toggle "Revert to Legacy Outlook."

### 6. Run first sync manually

```bash
cd ~/calendarBridge
chmod +x full_sync.sh
./full_sync.sh
```

This will:
- Export calendar from Outlook via AppleScript
- Open a browser window for Google OAuth (first time only)
- Sync events to Google Calendar

Check the output for "SYNC OK".

### 7. Install LaunchAgent

```bash
# Remove any old agents from previous versions
launchctl bootout gui/$(id -u)/net.leathermans.calendarbridge 2>/dev/null || true
launchctl bootout gui/$(id -u)/net.leathermans.calendarbridge.maintenance 2>/dev/null || true
launchctl bootout gui/$(id -u)/com.calendarbridge.sync 2>/dev/null || true

# Install new agent
cp ~/calendarBridge/launchd/com.calendarbridge.sync.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.calendarbridge.sync.plist
```

### 8. Verify it's running

```bash
launchctl list | grep calendarbridge
# Should show: -  0  com.calendarbridge.sync
```

### 9. Disable on old Mac

On the OLD MacBook, stop the LaunchAgent so only one machine syncs:

```bash
launchctl bootout gui/$(id -u)/net.leathermans.calendarbridge 2>/dev/null || true
launchctl bootout gui/$(id -u)/net.leathermans.calendarbridge.maintenance 2>/dev/null || true
launchctl bootout gui/$(id -u)/com.calendarbridge.sync 2>/dev/null || true
```

## Verify Everything Works

```bash
# Check last sync
tail -20 $(ls -t ~/calendarBridge/logs/full_sync_*.log | head -1)

# Check state
python3 -c "
import json
s = json.load(open('$HOME/calendarBridge/sync_state.json'))
print(f\"Events tracked: {len(s.get('events', {}))}\")
print(f\"Last sync: {s.get('last_sync', 'never')}\")
"

# Run health check
chmod +x ~/calendarBridge/maintenance.sh
~/calendarBridge/maintenance.sh
```

## What Changed in v7.0.0

1. **ICS staleness guard** — refuses to sync data older than 2 hours (configurable via `max_ics_age_hours`). Prevents stale data from deleting recent events.
2. **Atomic state writes** — sync_state.json can't be corrupted by crashes or power loss.
3. **Single LaunchAgent** — `com.calendarbridge.sync` replaces all previous agents.
4. **Pinned dependencies** — all Python packages version-locked.
5. **Exponential backoff** — rate limits (429) and server errors (5xx) retry with backoff instead of failing.
6. **Merged wrapper** — `full_sync.sh` handles everything (no more `run_sync_wrapper.sh`).
7. **Config validation** — typos in `calendar_config.json` caught immediately with clear errors.
8. **macOS notifications** — alerts on sync failures and significant changes.
9. **Streamlined logging** — single log path, no duplicate output.
10. **Removed dead code** — `clean_ics_files.py` deprecated (unused since v6.1).

## Troubleshooting

**"ICS file is X hours old"** — Outlook wasn't running when the sync fired. Open Outlook and run `./full_sync.sh` manually.

**"Config validation errors"** — Check `calendar_config.json` for typos. The error message tells you exactly what's wrong.

**No macOS notifications** — Check System Settings → Notifications → Script Editor is allowed.

**Token expired** — Delete `token.json` and run `./full_sync.sh` to re-authenticate.
