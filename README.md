# Mac Outlook → Google Calendar Sync (v3.7)

This project provides a robust, one-way synchronization from a local Microsoft Outlook for macOS client to a Google Calendar. It runs automatically in the background on an hourly schedule using macOS's native `launchd` service.

---

## Key Features

- **Automated Hourly Sync:** Runs every hour via a `launchd` agent.
- **Robust Event Handling:**
    - Expands recurring events to ensure all instances are synced.
    - Intelligently detects and converts Outlook's "timed" all-day events (e.g., "12a - 12a") into true, banner-style all-day events for proper display on Google Calendar and iOS.
    - Creates, updates, and deletes events in Google Calendar to maintain a perfect mirror of the Outlook source.
- **Intelligent & Resilient:**
    - Generates its own stable, unique IDs for each event to prevent issues with incompatible Outlook UIDs.
    - Handles intermittent Outlook `AppleEvent timed out` errors gracefully by succeeding on the next hourly run.
    - Manages API rate limits with an automatic backoff mechanism.
- **Safe & Private:** All authentication tokens and calendar data are stored and processed locally on your Mac.

## Repository Layout

The project has been simplified to a few core files:

```
calendarBridge/
├── calendar_config.json
├── calendar_sync.py
├── cleanup_synced_events.py
├── exportEvents.scpt
├── requirements.txt
├── sync.sh
├── net.leathermans.calendarbridge.plist
└── .gitignore
```

---

## Setup Instructions

### 1. Initial Setup

```bash
# Navigate to your project directory
cd /Users/bel/calendarBridge

# Create a Python virtual environment
python3 -m venv .venv

# Activate the environment
source .venv/bin/activate

# Install required packages
pip install -r requirements.txt
```

### 2. Google API Credentials

Ensure you have a `credentials.json` file from your Google Cloud Platform project placed in the `calendarBridge` directory. The first time you run the sync, it will open a browser window for you to authorize the application, which will create a `token.json` file.

### 3. Configuration

Edit the `calendar_config.json` file to match your setup. The script will look for the Outlook calendar by its name and index.

```json
{
    "outlook_calendar_name": "Calendar",
    "outlook_calendar_index": 2,
    "google_calendar_id": "your-email@gmail.com",
    "sync_days_past": 90,
    "sync_days_future": 120,
    "api_delay_seconds": 0.1
}
```

### 4. Automation with `launchd`

To run the sync automatically every hour:

```bash
# Move the service file to your user's LaunchAgents directory
mv net.leathermans.calendarbridge.plist ~/Library/LaunchAgents/

# Load and start the service
launchctl load ~/Library/LaunchAgents/net.leathermans.calendarbridge.plist
```

You can monitor the logs to see the sync happen: `tail -f logs/calendar_sync.log`
