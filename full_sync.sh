#!/usr/bin/env bash
# ============================================================================
# CalendarBridge Full Sync v7.0.0
#
# Single entry point for both LaunchAgent and manual runs.
# Merges the old run_sync_wrapper.sh logic — no separate wrapper needed.
# ============================================================================
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

notify() {
    osascript -e "display notification \"$1\" with title \"CalendarBridge\"" 2>/dev/null || true
}

cleanup() {
    rm -f "${LOCK_FILE}"
}
trap cleanup EXIT

# ============================================================================
# Lock: prevent concurrent syncs
# ============================================================================
if [ -f "${LOCK_FILE}" ]; then
    OLD_PID=$(cat "${LOCK_FILE}" 2>/dev/null)
    if [ -n "${OLD_PID}" ] && kill -0 "${OLD_PID}" 2>/dev/null; then
        log "Another sync running (PID ${OLD_PID}), exiting"
        exit 0
    fi
    rm -f "${LOCK_FILE}"
fi
echo $$ > "${LOCK_FILE}"

log "========== CalendarBridge v7.0.0 :: ${LOG_TS} =========="

# ============================================================================
# Pre-flight checks
# ============================================================================
if [ ! -x "${PY}" ]; then
    log "ERROR: Python not found at ${PY}"
    notify "Python venv missing — run setup"
    exit 1
fi

if [ ! -f "${ROOT}/credentials.json" ]; then
    log "ERROR: credentials.json not found"
    notify "Google credentials missing"
    exit 1
fi

# ============================================================================
# Step 1: Check Outlook (merged from run_sync_wrapper.sh)
# ============================================================================
log "[1/3] Checking Outlook..."
OUTLOOK_RUNNING=0
if osascript -e 'tell application "System Events" to (name of processes) contains "Microsoft Outlook"' 2>/dev/null | grep -q "true"; then
    OUTLOOK_RUNNING=1
fi

if [ "${OUTLOOK_RUNNING}" -eq 0 ]; then
    log "Outlook not running — checking cached ICS freshness"
    ICS_FILE="${ROOT}/outbox/outlook_full_export.ics"
    if [ -f "${ICS_FILE}" ]; then
        # Let safe_sync.py handle the staleness check — it will abort if too old
        log "Using cached ICS (staleness guard will validate)"
    else
        log "No cached ICS and Outlook not running — nothing to sync"
        exit 0
    fi
    SKIP_EXPORT=1
else
    SKIP_EXPORT=0
fi

# ============================================================================
# Step 2: Export from Outlook
# ============================================================================
if [ "${SKIP_EXPORT}" -eq 0 ]; then
    log "[2/3] Exporting from Outlook..."
    CAL_NAME=$("${PY}" -c "
import json
c = json.load(open('${ROOT}/calendar_config.json'))
print(c.get('outlook_calendar_name', 'Calendar'))
" 2>/dev/null || echo "Calendar")

    CAL_INDEX=$("${PY}" -c "
import json
c = json.load(open('${ROOT}/calendar_config.json'))
print(c.get('outlook_calendar_index', 2))
" 2>/dev/null || echo "2")

    if osascript "${ROOT}/exportEvents.scpt" "${CAL_NAME}" "${CAL_INDEX}" >> "${LOG_FILE}" 2>&1; then
        log "Export OK"
    else
        log "Export failed"
        # Check if cached ICS exists and is recent enough to be useful
        ICS_FILE="${ROOT}/outbox/outlook_full_export.ics"
        if [ -f "${ICS_FILE}" ]; then
            log "Will attempt sync with cached ICS (staleness guard applies)"
        else
            log "ERROR: Export failed and no cached ICS available"
            notify "Calendar export failed — no data to sync"
            exit 1
        fi
    fi
else
    log "[2/3] Skipping export (Outlook not running)"
fi

# ============================================================================
# Step 3: Run sync
# ============================================================================
ICS_FILE="${ROOT}/outbox/outlook_full_export.ics"
if [ ! -f "${ICS_FILE}" ]; then
    log "ERROR: No ICS file at ${ICS_FILE}"
    notify "No calendar data to sync"
    exit 1
fi

log "[3/3] Running sync..."
"${PY}" "${ROOT}/safe_sync.py" 2>&1 | tee -a "${LOG_FILE}"
SYNC_EXIT=${PIPESTATUS[0]}

if [ ${SYNC_EXIT} -eq 0 ]; then
    log "SYNC OK"
else
    log "SYNC FAILED (exit ${SYNC_EXIT})"
    # Notification is handled by safe_sync.py itself for detailed errors.
    # Only notify here for unexpected crashes.
    if [ ${SYNC_EXIT} -gt 1 ]; then
        notify "Sync crashed (exit ${SYNC_EXIT}) — check logs"
    fi
fi

# ============================================================================
# Housekeeping: clean old logs
# ============================================================================
find "${LOG_DIR}" -name "full_sync_*.log" -type f -mtime +7 -delete 2>/dev/null

exit ${SYNC_EXIT}
