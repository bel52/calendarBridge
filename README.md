# CalendarBridge (Outlook → Google Calendar)

**Version:** 7.0.2  
**Repo:** `github.com/bel52/calendarBridge`

Automated one-way sync from Microsoft Outlook (macOS Legacy mode) to Google Calendar. Runs locally on your Mac — no servers, no cloud dependencies.

---

## How It Works

1. AppleScript exports your Outlook calendar to ICS
2. Python parses ICS and compares against Google Calendar
3. Creates, updates, or deletes events to keep Google in sync
4. Tracks state locally to avoid duplicates and detect changes

---

## Key Features

- **Duplicate-safe**: Deterministic UID+start key matching prevents duplicate creation
- **Multi-VCALENDAR**: Handles complex Outlook exports with multiple calendar blocks
- **All-day event accuracy**: Proper formatting for iOS banner display
- **Recurring event expansion**: Syncs all instances within configured time window
- **Conservative deletion**: Only removes events created by CalendarBridge
- **Idempotent**: Run as often as needed — no side effects
- **Crash-safe**: Atomic state file writes via temp+rename
- **Staleness guard**: Refuses to sync ICS data older than 2 hours (configurable)
- **Retry with backoff**: Exponential backoff on Google API 429/5xx errors
- **Timezone-safe**: Key matching works correctly across timezone changes
- **macOS notifications**: Alerts on sync failures and significant changes

---

## Sync Schedule

CalendarBridge uses three layers to ensure reliability:

| Mechanism | Frequency | Purpose |
|-----------|-----------|---------|
| **LaunchAgent** | Every 30 min | Primary sync trigger (`com.calendarbridge.sync`) |
| **Cron fallback** | Every 30 min | Catches LaunchAgent failures after sleep/wake |
| **Maintenance cron** | Daily 8 AM | Health check, log cleanup, state backup, auto-recovery |

The LaunchAgent (`StartInterval`) can silently stop firing after macOS sleep/wake cycles. The cron fallback ensures syncs continue regardless. If maintenance detects no sync in 2+ hours, it reloads the agent and triggers a sync automatically.

---

## Prerequisites

- **macOS** with Microsoft Outlook in **Legacy mode**
- **Python 3.11+**
- **Google OAuth credentials** (`credentials.json`)

**Important:** New Outlook for Mac does not support AppleScript. You must use Legacy Outlook (supported by Microsoft until November 2026).

---

## Quick Start

See [DEPLOY.md](DEPLOY.md) for full step-by-step deployment on a new Mac.

```bash
cd ~
git clone https://github.com/bel52/calendarBridge.git
cd calendarBridge
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
# Add credentials.json, then:
./full_sync.sh
```

---

## Configuration

**File:** `calendar_config.json`

| Key | Type | Default | Description |
|-----|------|---------|-------------|
| `google_calendar_id` | string | *(required)* | Target Google Calendar ID |
| `outlook_calendar_name` | string | `Calendar` | Outlook calendar name to export |
| `outlook_calendar_index` | int | `2` | Index when multiple calendars share a name |
| `timezone` | string | *(required)* | IANA timezone (e.g., `America/New_York`) |
| `sync_days_past` | int | *(required)* | Days of history to sync (1–365) |
| `sync_days_future` | int | *(required)* | Days ahead to sync (1–365) |
| `api_delay_seconds` | float | `1.05` | Delay between Google API calls |
| `max_ics_age_hours` | float | `2.0` | Max ICS file age before refusing to sync |
| `enable_notifications` | bool | `true` | macOS notifications on changes/failures |

---

## File Structure

```
~/calendarBridge/
├── safe_sync.py              # Core sync engine
├── full_sync.sh              # Entry point (export + sync)
├── maintenance.sh            # Health check + log cleanup + auto-recovery
├── cleanup_duplicates.py     # Duplicate removal tool (manual use)
├── exportEvents.scpt         # AppleScript Outlook exporter
├── calendar_config.json      # Sync configuration
├── credentials.json          # Google OAuth credentials (not in git)
├── token.json                # Google auth token (not in git)
├── sync_state.json           # Event tracking state (not in git)
├── requirements.txt          # Python dependencies (pinned)
├── .venv/                    # Python virtual environment
├── outbox/                   # Exported ICS files
├── logs/                     # Sync and maintenance logs
│   ├── full_sync_*.log       # Per-sync logs (14-day retention)
│   ├── cron-sync.log         # Cron fallback output
│   ├── maintenance.log       # Health check history
│   └── maintenance-cron.log  # Cron maintenance output
└── launchd/
    └── com.calendarbridge.sync.plist
```

---

## Commands

```bash
# Manual sync
cd ~/calendarBridge && ./full_sync.sh

# Check last sync
tail -10 $(ls -t ~/calendarBridge/logs/full_sync_*.log | head -1)

# Run health check
./maintenance.sh

# Check if LaunchAgent is loaded
launchctl list | grep calendarbridge

# Reload LaunchAgent (if stalled)
launchctl bootout gui/$(id -u) ~/Library/LaunchAgents/com.calendarbridge.sync.plist 2>/dev/null
launchctl bootstrap gui/$(id -u) ~/Library/LaunchAgents/com.calendarbridge.sync.plist

# Check cron jobs
crontab -l

# Dry-run duplicate cleanup
python3 cleanup_duplicates.py

# Apply duplicate cleanup
python3 cleanup_duplicates.py --apply
```

---

## Troubleshooting

**"ICS file is X hours old"** — Outlook wasn't running when the sync fired. Open Outlook and run `./full_sync.sh` manually.

**LaunchAgent not firing** — Common after sleep/wake. The cron fallback handles this automatically. To fix immediately: reload with `launchctl bootout`/`bootstrap` commands above.

**Token expired** — Delete `token.json` and run `./full_sync.sh` to re-authenticate via browser.

**Config validation errors** — Check `calendar_config.json` for typos. Error messages are specific.

**No macOS notifications** — Check System Settings → Notifications → Script Editor is allowed.

---

## Changelog

### v7.0.2 (2026-04-03)
- Added cron fallback sync (every 30 min) to survive LaunchAgent sleep/wake failures
- Maintenance auto-recovery: reloads agent and triggers sync when stale >2 hours (was warn-only)
- Staleness threshold lowered from 24h to 2h for faster detection

### v7.0.1 (2026-03-20)
- Timezone-safe key matching (Google API returns times in user's current timezone; keys now normalized to config timezone before comparison)
- Fixed duplicate crisis caused by timezone key mismatch (~9,000 duplicates cleaned)

### v7.0.0 (2026-03-20)
- Atomic state file writes (crash-safe via temp+rename)
- ICS staleness guard (configurable max age, default 2h)
- Exponential backoff on Google API rate limits (429) and server errors (5xx)
- Config validation on startup with clear error messages
- macOS notifications on sync failures and significant changes
- Single LaunchAgent replaces all previous agents
- Pinned Python dependencies
- Merged wrapper — `full_sync.sh` handles everything
- Removed deprecated `clean_ics_files.py`
