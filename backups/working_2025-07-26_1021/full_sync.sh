#!/usr/bin/env bash
# Master pipeline: export → clean → sync
set -euo pipefail

LOG_DIR="$HOME/calendarBridge/logs"
mkdir -p "$LOG_DIR"
LOG="$LOG_DIR/$(date +%F_%H%M).log"

{
  echo "=== $(date) ==="
  osascript "$HOME/calendarBridge/exportEvents.scpt"
  python3   "$HOME/calendarBridge/clean_ics_files.py"
  python3   "$HOME/calendarBridge/safe_sync.py"
  echo "Success."
} | tee -a "$LOG"
