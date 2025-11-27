#!/usr/bin/env bash
set -o pipefail

ROOT="${HOME}/calendarBridge"
VENV="${ROOT}/.venv"
PY="${VENV}/bin/python"
LOG_DIR="${ROOT}/logs"
LOCK_FILE="${ROOT}/.sync.lock"
LOG_TS="$(date '+%Y-%m-%d_%H-%M-%S')"
LOG_FILE="${LOG_DIR}/full_sync_${LOG_TS}.log"

mkdir -p "${LOG_DIR}"

log() {
    echo "[$(date '+%H:%M:%S')] $*" | tee -a "${LOG_FILE}"
}

cleanup() {
    rm -f "${LOCK_FILE}"
}
trap cleanup EXIT

if [ -f "${LOCK_FILE}" ]; then
    OLD_PID=$(cat "${LOCK_FILE}" 2>/dev/null)
    if [ -n "${OLD_PID}" ] && kill -0 "${OLD_PID}" 2>/dev/null; then
        log "Another sync running (PID ${OLD_PID}), exiting"
        exit 0
    fi
    rm -f "${LOCK_FILE}"
fi

echo $$ > "${LOCK_FILE}"

log "========== CalendarBridge Full Sync :: ${LOG_TS} =========="

if [ ! -x "${PY}" ]; then
    log "ERROR: Python not found at ${PY}"
    exit 1
fi

export CALBRIDGE_QUOTA_USER="${CALBRIDGE_QUOTA_USER:-brett@leathermans.net}"

log "[STEP 1/3] Checking Outlook..."
if ! osascript -e 'tell application "System Events" to (name of processes) contains "Microsoft Outlook"' 2>/dev/null | grep -q "true"; then
    log "Outlook not running - using cached ICS"
    SKIP_EXPORT=1
else
    SKIP_EXPORT=0
fi

if [ "${SKIP_EXPORT}" -eq 0 ]; then
    log "[STEP 2/3] Exporting from Outlook..."
    CAL_NAME=$(python3 -c "import json; c=json.load(open('${ROOT}/calendar_config.json')); print(c.get('outlook_calendar_name', 'Calendar'))" 2>/dev/null || echo "Calendar")
    CAL_INDEX=$(python3 -c "import json; c=json.load(open('${ROOT}/calendar_config.json')); print(c.get('outlook_calendar_index', 2))" 2>/dev/null || echo "2")
    
    if osascript "${ROOT}/exportEvents.scpt" "${CAL_NAME}" "${CAL_INDEX}" >> "${LOG_FILE}" 2>&1; then
        log "Export OK"
    else
        log "Export failed, using cached data"
    fi
else
    log "[STEP 2/3] Skipping export"
fi

ICS_FILE="${ROOT}/outbox/outlook_full_export.ics"
if [ ! -f "${ICS_FILE}" ]; then
    log "ERROR: No ICS file"
    exit 1
fi

log "[STEP 3/3] Running sync..."
"${PY}" "${ROOT}/safe_sync.py" 2>&1 | tee -a "${LOG_FILE}"
SYNC_EXIT=${PIPESTATUS[0]}

if [ ${SYNC_EXIT} -eq 0 ]; then
    log "SYNC OK"
else
    log "SYNC FAILED (exit ${SYNC_EXIT})"
fi

find "${LOG_DIR}" -name "*.log" -type f -mtime +7 -delete 2>/dev/null

exit ${SYNC_EXIT}
