#!/usr/bin/env bash
set -euo pipefail

LOG_DIR="${HOME}/calendarBridge/logs"
mkdir -p "$LOG_DIR"

# Skip if Outlook isn't running (no GUI context = no AppleScript export)
if ! /usr/bin/osascript -e 'tell application "System Events" to (name of processes) contains "Microsoft Outlook"' >/dev/null 2>&1; then
  echo "[$(date)] Outlook not running; skipping sync." >> "$LOG_DIR/last_run.err"
  exit 0
fi

# Run the full pipeline
exec "${HOME}/calendarBridge/full_sync.sh" >> "$LOG_DIR/last_run.log" 2>> "$LOG_DIR/last_run.err"
