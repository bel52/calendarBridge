#!/usr/bin/env bash
# ============================================================================
# CalendarBridge Maintenance v7.0.1
# Runs weekly to clean logs, verify health, and backup state.
# ============================================================================
set -euo pipefail

ROOT="${HOME}/calendarBridge"
LOG_DIR="${ROOT}/logs"
HEALTH_LOG="${LOG_DIR}/maintenance.log"
AGENT_LABEL="com.calendarbridge.sync"

log() {
    echo "[$(date '+%Y-%m-%d %H:%M:%S')] $*" | tee -a "${HEALTH_LOG}"
}

# ============================================================================
# Log Cleanup
# ============================================================================

cleanup_logs() {
    log "=== Log Cleanup ==="

    # Remove sync logs older than 14 days
    find "${LOG_DIR}" -name "full_sync_*.log" -type f -mtime +14 -delete 2>/dev/null || true
    local sync_logs
    sync_logs=$(find "${LOG_DIR}" -name "full_sync_*.log" -type f | wc -l | tr -d ' ')
    log "Kept ${sync_logs} recent sync logs"

    # Clean LaunchAgent stdout/stderr logs
    for f in /tmp/calendarbridge.stdout.log /tmp/calendarbridge.stderr.log; do
        if [ -f "$f" ]; then
            local size
            size=$(stat -f%z "$f" 2>/dev/null || echo 0)
            if [ "${size}" -gt 5242880 ]; then
                > "$f"
                log "Truncated $(basename "$f") (was ${size} bytes)"
            fi
        fi
    done

    # Clean maintenance logs older than 30 days
    find "${LOG_DIR}" -name "maintenance.log.*" -type f -mtime +30 -delete 2>/dev/null || true

    log "Log cleanup complete"
}

# ============================================================================
# Health Checks
# ============================================================================

health_check() {
    log "=== Health Check ==="

    # Check LaunchAgent status
    if launchctl list 2>/dev/null | grep -q "${AGENT_LABEL}"; then
        log "OK: LaunchAgent running"
    else
        log "WARN: LaunchAgent NOT running — attempting load"
        launchctl load ~/Library/LaunchAgents/${AGENT_LABEL}.plist 2>&1 | tee -a "${HEALTH_LOG}" || true
    fi

    # Check last sync status
    local last_log
    last_log=$(ls -t "${LOG_DIR}"/full_sync_*.log 2>/dev/null | head -1)
    if [ -n "${last_log}" ]; then
        if grep -q "SYNC OK" "${last_log}"; then
            log "OK: Last sync succeeded"
        elif grep -q "SYNC COMPLETE" "${last_log}"; then
            log "OK: Last sync completed"
        else
            log "WARN: Last sync may have failed — check ${last_log}"
        fi

        # Check how old the last sync is
        local last_mtime
        last_mtime=$(stat -f%m "${last_log}" 2>/dev/null || echo 0)
        local now
        now=$(date +%s)
        local age_hours=$(( (now - last_mtime) / 3600 ))
        if [ "${age_hours}" -gt 24 ]; then
            log "WARN: Last sync was ${age_hours}h ago"
        fi
    else
        log "WARN: No sync logs found"
    fi

    # Check state file
    if [ -f "${ROOT}/sync_state.json" ]; then
        local event_count
        event_count=$(python3 -c "
import json
print(len(json.load(open('${ROOT}/sync_state.json')).get('events', {})))
" 2>/dev/null || echo "unknown")
        log "OK: State file has ${event_count} events tracked"
    else
        log "WARN: State file missing"
    fi

    # Check credentials
    if [ -f "${ROOT}/credentials.json" ]; then
        log "OK: Google credentials present"
    else
        log "ERROR: credentials.json missing"
    fi

    if [ -f "${ROOT}/token.json" ]; then
        log "OK: Google token present"
    else
        log "WARN: token.json missing — re-auth needed on next sync"
    fi

    # Check venv
    if [ -x "${ROOT}/.venv/bin/python" ]; then
        log "OK: Python venv intact"
    else
        log "ERROR: Python venv broken or missing"
    fi

    # Disk usage
    local disk_usage
    disk_usage=$(du -sh "${ROOT}" 2>/dev/null | awk '{print $1}')
    log "INFO: Disk usage: ${disk_usage}"

    log "Health check complete"
}

# ============================================================================
# State File Backup
# ============================================================================

backup_state() {
    log "=== State Backup ==="

    if [ -f "${ROOT}/sync_state.json" ]; then
        local backup_name="sync_state.$(date +%Y%m%d).json"
        cp "${ROOT}/sync_state.json" "${LOG_DIR}/${backup_name}"

        # Remove backups older than 7 days
        find "${LOG_DIR}" -name "sync_state.*.json" -type f -mtime +7 -delete 2>/dev/null || true

        local backup_count
        backup_count=$(find "${LOG_DIR}" -name "sync_state.*.json" -type f | wc -l | tr -d ' ')
        log "State backed up (${backup_count} backups kept)"
    else
        log "WARN: No state file to backup"
    fi
}

# ============================================================================
# Main
# ============================================================================

main() {
    # Rotate maintenance log if over 1MB
    if [ -f "${HEALTH_LOG}" ]; then
        local size
        size=$(stat -f%z "${HEALTH_LOG}" 2>/dev/null || echo 0)
        if [ "${size}" -gt 1048576 ]; then
            mv "${HEALTH_LOG}" "${HEALTH_LOG}.$(date +%Y%m%d)"
        fi
    fi

    log "========================================"
    log "CalendarBridge Maintenance v7.0.1"
    log "========================================"

    cleanup_logs
    echo "" >> "${HEALTH_LOG}"

    backup_state
    echo "" >> "${HEALTH_LOG}"

    health_check
    echo "" >> "${HEALTH_LOG}"

    log "========================================"
    log "Maintenance Complete"
    log "========================================"
}

main "$@"
