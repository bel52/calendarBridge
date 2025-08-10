#!/bin/bash
set -euo pipefail
export PYTHONDONTWRITEBYTECODE=1
LOG_DIR="$HOME/calendarBridge/logs"
mkdir -p "$LOG_DIR"
LOG="$LOG_DIR/$(date +%F_%H%M).log"

# Ensure Outlook has UI context
osascript <<END
tell application "Microsoft Outlook"
    activate
end tell
END

VENV_PY="$HOME/calendarBridge/.venv/bin/python"

{
  echo "=== $(date) ==="

  osascript   "$HOME/calendarBridge/exportEvents.scpt"
  "$VENV_PY"  "$HOME/calendarBridge/clean_ics_files.py"
  "$VENV_PY"  "$HOME/calendarBridge/safe_sync.py"

  echo "=== DONE $(date) ==="
} 2>&1 | tee -a "$LOG"
