#!/usr/bin/env bash
set -euo pipefail

# --- CONFIG ---
ROOT="${HOME}/calendarBridge"
VENV="${ROOT}/.venv"
OUTBOX="${ROOT}/outbox"
LOGDIR="${ROOT}/logs"
PY=${VENV}/bin/python
DATESTAMP="$(date '+%Y-%m-%d_%H-%M-%S')"
LOGFILE="${LOGDIR}/full_sync_${DATESTAMP}.log"

# Pass these to select the right Outlook calendar:
CALENDAR_NAME="${CALENDAR_NAME:-Calendar}"
CALENDAR_INDEX="${CALENDAR_INDEX:-2}"   # 2nd calendar named "Calendar"

mkdir -p "${LOGDIR}" "${OUTBOX}"

exec > >(tee -a "${LOGFILE}") 2>&1

echo "========== CalendarBridge Full Sync :: ${DATESTAMP} =========="
echo "[INFO] Using ROOT=${ROOT}"
echo "[INFO] Python path: ${PY}"

# shellcheck disable=SC1091
source "${VENV}/bin/activate" || true
echo "[INFO] Python version: $(python -V 2>&1 || echo 'N/A')"
echo "[INFO] PIP packages:"
pip freeze | sed 's/^/[PKG] /' || true

echo "[STEP] 1/4 Export Outlook events -> ${OUTBOX}"
echo "[INFO] Export args: name='${CALENDAR_NAME}' index='${CALENDAR_INDEX}'"
osascript "${ROOT}/exportEvents.scpt" "${CALENDAR_NAME}" "${CALENDAR_INDEX}" || { echo "[ERR ] AppleScript export failed"; exit 10; }

echo "[STEP] 2/4 Count .ics files"
ICS_COUNT=$(find "${OUTBOX}" -type f -name '*.ics' | wc -l | tr -d ' ')
echo "[INFO] Found ${ICS_COUNT} .ics files in outbox"
if [[ "${ICS_COUNT}" -eq 0 ]]; then
  echo "[WARN] No .ics exported. Stopping so we don't wipe Google with empty input."
  exit 20
fi

echo "[STEP] 3/4 Clean headers if cleaner exists"
if [[ -f "${ROOT}/clean_ics_files.py" ]]; then
  python "${ROOT}/clean_ics_files.py" || { echo "[WARN] clean_ics_files.py had issues, continuing"; }
fi

echo "[STEP] 4/4 Run main sync"
if [[ -f "${ROOT}/calendar_sync.py" ]]; then
  python "${ROOT}/calendar_sync.py" --verbose || { echo "[ERR ] calendar_sync.py failed"; exit 30; }
elif [[ -f "${ROOT}/safe_sync.py" ]]; then
  python "${ROOT}/safe_sync.py" --verbose || { echo "[ERR ] safe_sync.py failed"; exit 31; }
else
  echo "[ERR ] No main sync script (calendar_sync.py or safe_sync.py) present"
  exit 32
fi

echo "[DONE] Full sync completed at ${DATESTAMP}"
