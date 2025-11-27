# CalendarBridge (Outlook → Google Calendar)

**Version 6.1.0** - Production-ready sync from Microsoft Outlook (macOS) to Google Calendar

A reliable, automated calendar synchronization tool that runs locally on your Mac. No servers, no cloud dependencies—your credentials stay on your machine.

---

## Key Features

- **Duplicate-Safe**: Deterministic event matching by UID+start time prevents duplicate creation
- **Multi-VCALENDAR Support**: Handles complex Outlook exports with multiple calendar blocks
- **All-Day Event Accuracy**: Proper formatting for iOS/iPhone banner display
- **Recurring Event Expansion**: Syncs all instances within your configured time window
- **Conservative Deletion**: Only removes events created by CalendarBridge
- **Idempotent**: Run as often as needed—no duplicate creation or data loss
- **Automated**: Runs every 30 minutes via macOS LaunchAgent

---

## Contents

- [Prerequisites](#prerequisites)
- [Installation](#installation)
- [Configuration](#configuration)
- [Manual Sync](#manual-sync)
- [Automated Sync (LaunchAgent)](#automated-sync-launchagent)
- [Troubleshooting](#troubleshooting)
- [Maintenance](#maintenance)

---

## Prerequisites

- **macOS** with Microsoft Outlook app installed
- **Python 3.11+** (tested with 3.14)
- **Google OAuth credentials** (`credentials.json`)
- Active Google Calendar account

---

## Installation

### 1. Clone Repository

```bash
cd ~
git clone https://github.com/bel52/calendarBridge.git
cd calendarBridge
```

### 2. Create Virtual Environment

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install --upgrade pip
pip install -r requirements.txt
```

### 3. Add Google OAuth Credentials

Place your `credentials.json` file in the `calendarBridge` directory. On first run, a browser window will open to authorize access, creating `token.json` automatically.

---

## Configuration

Edit `calendar_config.json`:

```json
{
  "outlook_calendar_name": "Calendar",
  "outlook_calendar_index": 2,
  "google_calendar_id": "your-email@example.com",
  "sync_days_past": 60,
  "sync_days_future": 90,
  "api_delay_seconds": 1.05,
  "timezone": "America/New_York"
}
```

**Parameters:**

- `outlook_calendar_name`: Name of Outlook calendar to sync
- `outlook_calendar_index`: If multiple calendars share the same name, specify which one (1-based)
- `google_calendar_id`: Target Google Calendar ID (usually your email)
- `sync_days_past`: Include events from X days in the past
- `sync_days_future`: Include events up to Y days in the future
- `api_delay_seconds`: Delay between API calls (increase if hitting rate limits)
- `timezone`: Your local timezone (IANA format)

**Note:** Recurring events that start outside the time window will still sync if any instance falls within it.

---

## Manual Sync

To run a one-time sync:

```bash
cd ~/calendarBridge
source .venv/bin/activate
./full_sync.sh
```

This will:
1. Export events from Outlook via AppleScript
2. Parse and clean the ICS file
3. Sync to Google Calendar (create, update, or delete as needed)

Logs are written to `logs/full_sync_YYYY-MM-DD_HH-MM-SS.log`.

---

## Automated Sync (LaunchAgent)

### Setup

The LaunchAgent is already configured at:
```
~/Library/LaunchAgents/net.leathermans.calendarbridge.plist
```

It runs every 30 minutes (`StartInterval: 1800`) and on login (`RunAtLoad: true`).

### Load LaunchAgent

```bash
launchctl load ~/Library/LaunchAgents/net.leathermans.calendarbridge.plist
launchctl kickstart -k gui/$(id -u)/net.leathermans.calendarbridge
```

### Check Status

```bash
launchctl list | grep calendarbridge
tail -f ~/calendarBridge/logs/launchd.out
```

### Unload LaunchAgent

```bash
launchctl bootout gui/$(id -u)/net.leathermans.calendarbridge
```

**Note:** LaunchAgents don't run while the Mac is asleep. They resume on wake.

---

## Project Structure

```
calendarBridge/
├── calendar_config.json        # Sync configuration
├── credentials.json            # Google OAuth client (not committed)
├── token.json                  # Google access token (auto-generated)
├── exportEvents.scpt           # AppleScript to export from Outlook
├── clean_ics_files.py          # ICS pre-processor (unused in v6.1+)
├── safe_sync.py                # Main sync engine
├── cleanup_duplicates.py       # Duplicate detection/removal tool
├── full_sync.sh                # Orchestrates export → sync pipeline
├── run_sync_wrapper.sh         # LaunchAgent wrapper (checks Outlook)
├── sync_state.json             # Tracks synced events (auto-managed)
├── outbox/                     # Temporary ICS export storage
├── logs/                       # Sync logs and LaunchAgent output
├── .venv/                      # Python virtual environment
└── README.md                   # This file
```

---

## Troubleshooting

### Sync Failures

**Symptoms:** `SYNC FAILED (exit 1)` in logs

**Solutions:**
1. Check `logs/full_sync_*.log` for specific error
2. Verify Outlook is running (LaunchAgent skips if not)
3. Ensure `credentials.json` and `token.json` are valid
4. Check network connectivity

### Duplicate Events

**Prevention:** Version 6.1.0 uses deterministic UID+start matching to prevent duplicates.

**Cleanup:** If duplicates exist from earlier versions:
```bash
cd ~/calendarBridge
source .venv/bin/activate
./cleanup_duplicates.py          # Dry-run
./cleanup_duplicates.py --apply  # Delete duplicates
```

### All-Day Events Display Incorrectly

Version 6.1.0 detects all-day events using multiple heuristics:
- `X-MICROSOFT-CDO-ALLDAYEVENT:TRUE` header
- `VALUE=DATE` format
- Midnight-to-midnight time spans

Events are created with `transparency: transparent` for proper iOS display.

### Rate Limiting

**Symptoms:** `rateLimitExceeded` errors in logs

**Solutions:**
1. Increase `api_delay_seconds` in `calendar_config.json` (try 1.5 or 2.0)
2. Reduce sync window (`sync_days_past`/`sync_days_future`)
3. Wait 10-15 minutes and retry

### Re-authentication

If Google tokens expire or become invalid:
```bash
cd ~/calendarBridge
rm token.json
./full_sync.sh  # Will re-open browser for authorization
```

---

## Maintenance

### View Recent Logs

```bash
cd ~/calendarBridge
ls -lt logs/full_sync_*.log | head -5
tail -50 $(ls -t logs/full_sync_*.log | head -1)
```

### Check Sync State

```bash
python3 -c "
import json
with open('sync_state.json') as f:
    state = json.load(f)
    print(f'Tracked events: {len(state.get(\"events\", {}))}')
"
```

### Clean Up Old Logs

Logs older than 7 days are automatically deleted by `full_sync.sh`.

Manual cleanup:
```bash
find ~/calendarBridge/logs -name "*.log" -mtime +7 -delete
```

### Update CalendarBridge

```bash
cd ~/calendarBridge
git pull
source .venv/bin/activate
pip install --upgrade -r requirements.txt
launchctl kickstart -k gui/$(id -u)/net.leathermans.calendarbridge
```

---

## Version History

### 6.1.0 (2025-11-27)
- **Fixed:** Timezone key matching to prevent duplicate event creation
- **Added:** Multi-VCALENDAR block support for complex Outlook exports
- **Improved:** All-day event detection with multiple heuristics
- **Removed:** Deprecated diagnostic and test scripts
- **Production Ready:** Stable duplicate-safe sync

### 6.0.0 (2025-11-26)
- Complete rewrite with deterministic event IDs
- State-based change tracking
- Conservative deletion (CalendarBridge events only)

### 5.0.0 and earlier
- Initial implementations with various sync strategies

---

## Security & Privacy

- All credentials (`credentials.json`, `token.json`) remain on your Mac
- No data is sent to external servers except Google Calendar API
- OAuth tokens are stored locally and never committed to Git
- `.gitignore` prevents accidental credential commits

---

## Contributing

This is a personal project, but issues and PRs are welcome:
- Report bugs via GitHub Issues
- Include relevant log snippets
- Test changes with `./full_sync.sh` before submitting PRs

---

## License

MIT License - See repository for details

---

## Support

For issues or questions:
1. Check [Troubleshooting](#troubleshooting) section
2. Review recent logs in `logs/` directory
3. Open a GitHub issue with log excerpts

**Maintained by:** Brett Leatherman  
**Repository:** https://github.com/bel52/calendarBridge
