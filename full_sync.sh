#!/usr/bin/env bash
set -euo pipefail

ROOT="${HOME}/calendarBridge"
VENV="${ROOT}/.venv"
PY="${VENV}/bin/python"
LOG_TS="$(date '+%Y-%m-%d_%H-%M-%S')"

echo "========== CalendarBridge Full Sync :: ${LOG_TS} =========="
echo "[INFO] Using ROOT=${ROOT}"
echo "[INFO] Python path: ${PY}"
"${PY}" -V

# Light jitter so multiple launches don't align to the same second
sleep "$(python3 - <<'PY'
import random; print(round(random.uniform(0.3,0.7),2))
PY
)"

# Export env knobs read by safe_sync.py (if present)
export CALBRIDGE_QUOTA_USER="brett@leathermans.net"   # helps Google bucket requests
export CALBRIDGE_SLOW_START_MS=300                     # tiny initial pause before first write

echo "[STEP] 1/4 Export Outlook events -> ${ROOT}/outbox"
echo "[INFO] Export args: name='Calendar' index='2'"
osascript "${ROOT}/exportEvents.scpt" "Calendar" "2"

echo "[OK  ] Outlook export completed"

echo "[STEP] 2/4 Count .ics files"
count=$(find "${ROOT}/outbox" -type f -name '*.ics' | wc -l | tr -d ' ')
echo "[INFO] Found ${count} .ics files in outbox"

echo "[STEP] 3/4 Clean ICS headers"
"${PY}" "${ROOT}/clean_ics_files.py" --inbox "${ROOT}/outbox"

echo "[STEP] 4/4 Run main sync"
"${PY}" "${ROOT}/safe_sync.py" --config "${ROOT}/calendar_config.json"

echo "SYNC OK :: ${LOG_TS}"
