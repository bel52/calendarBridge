#!/usr/bin/env bash
set -euo pipefail

ICS_URL="${1:-}"
TARGET_CAL="${2:-Calendar 2}"
CATEGORY="${3:-Imported: Detroit Lions}"
JSON_OUT="/tmp/public_events.json"

if [[ -z "${ICS_URL}" ]]; then
  echo "Usage: $0 <ICS_URL> [Outlook Calendar Name] [Category Name]" >&2
  exit 1
fi

# Use your venv Python if available
if [[ -f "$HOME/calendarBridge/.venv/bin/python3" ]]; then
  PY="$HOME/calendarBridge/.venv/bin/python3"
else
  PY="python3"
fi

echo "[*] Fetching and parsing ICS…"
"$PY" "$HOME/calendarBridge/fetch_public_ics.py" "$ICS_URL" "$JSON_OUT"

echo "[*] Writing events into Outlook calendar: ${TARGET_CAL}"
osascript -l JavaScript "$HOME/calendarBridge/outlook_write_events.js" "$JSON_OUT" "$TARGET_CAL" "$CATEGORY"
