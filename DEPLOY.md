# CalendarBridge v7.0.2 — New Mac Deployment

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

### 4. Create directories

```bash
mkdir -p ~/calendarBridge/outbox
mkdir -p ~/calendarBridge/logs
```

### 5. Verify Outlook is in Legacy mode

Open Outlook. If it's in New Outlook, toggle "Revert to Legacy Outlook."

### 6. Run first sync manually

```bash
cd ~/calendarBridge
chmod +x full_sync.sh maintenance.sh
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
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.calendarbridge.sync.plist
```

### 8. Verify it's running

```bash
launchctl list | grep calendarbridge
# Should show a PID and com.calendarbridge.sync
```

### 9. Install cron jobs

The LaunchAgent can silently stop firing after macOS sleep/wake cycles. Cron provides a reliable fallback:

```bash
(crontab -l 2>/dev/null; echo "*/30 * * * * /Users/$(whoami)/calendarBridge/full_sync.sh >> /Users/$(whoami)/calendarBridge/logs/cron-sync.log 2>&1"; echo "0 8 * * * /Users/$(whoami)/calendarBridge/maintenance.sh >> /Users/$(whoami)/calendarBridge/logs/maintenance-cron.log 2>&1") | crontab -
crontab -l
```

You should see two entries:
- `*/30` — sync fallback every 30 minutes
- `0 8` — daily maintenance (health check, log cleanup, state backup, auto-recovery)

### 10. Disable on old Mac

On the OLD MacBook, stop everything so only one machine syncs:

```bash
launchctl bootout gui/$(id -u)/net.leathermans.calendarbridge 2>/dev/null || true
launchctl bootout gui/$(id -u)/net.leathermans.calendarbridge.maintenance 2>/dev/null || true
launchctl bootout gui/$(id -u)/com.calendarbridge.sync 2>/dev/null || true
crontab -r 2>/dev/null || true  # Remove all cron jobs (review first if you have others)
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
~/calendarBridge/maintenance.sh

# Verify cron
crontab -l
```

## Troubleshooting

**"ICS file is X hours old"** — Outlook wasn't running when the sync fired. Open Outlook and run `./full_sync.sh` manually.

**"Config validation errors"** — Check `calendar_config.json` for typos. The error message tells you exactly what's wrong.

**No macOS notifications** — Check System Settings → Notifications → Script Editor is allowed.

**Token expired** — Delete `token.json` and run `./full_sync.sh` to re-authenticate.

**LaunchAgent not firing after sleep** — This is expected macOS behavior. The cron fallback (step 9) handles it automatically. To fix immediately: `launchctl bootout gui/$(id -u) ~/Library/LaunchAgents/com.calendarbridge.sync.plist 2>/dev/null; launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.calendarbridge.sync.plist`
