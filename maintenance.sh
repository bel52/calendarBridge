#!/usr/bin/env bash
# CalendarBridge Maintenance Script
# Runs weekly to clean logs and verify health

set -euo pipefail

ROOT="${HOME}/calendarBridge"
LOG_DIR="${ROOT}/logs"
HEALTH_LOG="${LOG_DIR}/maintenance.log"

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
    local sync_logs=$(find "${LOG_DIR}" -name "full_sync_*.log" -type f | wc -l | tr -d ' ')
    log "Kept ${sync_logs} recent sync logs"
    
    # Rotate launchd logs if over 10MB
    if [ -f "${LOG_DIR}/launchd.out" ]; then
        local size=$(stat -f%z "${LOG_DIR}/launchd.out" 2>/dev/null || echo 0)
        if [ "${size}" -gt 10485760 ]; then
            mv "${LOG_DIR}/launchd.out" "${LOG_DIR}/launchd.out.$(date +%Y%m%d)"
            touch "${LOG_DIR}/launchd.out"
            log "Rotated launchd.out (was ${size} bytes)"
        fi
    fi
    
    # Keep only last 5 rotated launchd logs
    find "${LOG_DIR}" -name "launchd.out.*" -type f | sort -r | tail -n +6 | xargs rm -f 2>/dev/null || true
    
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
    if launchctl list | grep -q "net.leathermans.calendarbridge"; then
        log "✓ LaunchAgent running"
    else
        log "✗ LaunchAgent NOT running - attempting restart"
        launchctl load ~/Library/LaunchAgents/net.leathermans.calendarbridge.plist 2>&1 | tee -a "${HEALTH_LOG}"
    fi
    
    # Check last sync status
    local last_log=$(ls -t "${LOG_DIR}"/full_sync_*.log 2>/dev/null | head -1)
    if [ -n "${last_log}" ]; then
        if grep -q "SYNC OK" "${last_log}"; then
            local sync_time=$(grep "SYNC COMPLETE" "${last_log}" | tail -1 | awk '{print $NF}' | sed 's/s$//')
            log "✓ Last sync successful (${sync_time}s)"
        else
            log "✗ Last sync failed - check ${last_log}"
        fi
    else
        log "⚠ No sync logs found"
    fi
    
    # Check state file
    if [ -f "${ROOT}/sync_state.json" ]; then
        local event_count=$(python3 -c "import json; print(len(json.load(open('${ROOT}/sync_state.json')).get('events', {})))" 2>/dev/null || echo "unknown")
        log "✓ State file OK (${event_count} events tracked)"
    else
        log "✗ State file missing"
    fi
    
    # Check credentials
    if [ -f "${ROOT}/token.json" ]; then
        log "✓ Google credentials present"
    else
        log "✗ Google token missing - re-auth needed"
    fi
    
    # Check disk usage
    local disk_usage=$(du -sh "${ROOT}" 2>/dev/null | awk '{print $1}')
    log "ℹ Disk usage: ${disk_usage}"
    
    # Check outbox size
    local outbox_count=$(ls -1 "${ROOT}/outbox" 2>/dev/null | wc -l | tr -d ' ')
    log "ℹ Outbox files: ${outbox_count}"
    
    log "Health check complete"
}

# ============================================================================
# Outbox Cleanup
# ============================================================================

cleanup_outbox() {
    log "=== Outbox Cleanup ==="
    
    # Keep only the 3 most recent ICS exports
    cd "${ROOT}/outbox"
    ls -t outlook_full_export.ics.* 2>/dev/null | tail -n +4 | xargs rm -f 2>/dev/null || true
    
    # Clean old cleaned ICS files (keep only current)
    find . -name "clean_*.ics" -type f -mtime +1 -delete 2>/dev/null || true
    
    local count=$(ls -1 | wc -l | tr -d ' ')
    log "Outbox cleaned (${count} files remaining)"
}

# ============================================================================
# State File Backup
# ============================================================================

backup_state() {
    log "=== State Backup ==="
    
    if [ -f "${ROOT}/sync_state.json" ]; then
        # Keep daily backups for 7 days
        local backup_name="sync_state.$(date +%Y%m%d).json"
        cp "${ROOT}/sync_state.json" "${ROOT}/logs/${backup_name}"
        
        # Remove backups older than 7 days
        find "${LOG_DIR}" -name "sync_state.*.json" -type f -mtime +7 -delete 2>/dev/null || true
        
        local backup_count=$(find "${LOG_DIR}" -name "sync_state.*.json" -type f | wc -l | tr -d ' ')
        log "State backed up (${backup_count} backups kept)"
    else
        log "⚠ No state file to backup"
    fi
}

# ============================================================================
# Main
# ============================================================================

main() {
    # Rotate maintenance log if over 1MB
    if [ -f "${HEALTH_LOG}" ]; then
        local size=$(stat -f%z "${HEALTH_LOG}" 2>/dev/null || echo 0)
        if [ "${size}" -gt 1048576 ]; then
            mv "${HEALTH_LOG}" "${HEALTH_LOG}.$(date +%Y%m%d)"
        fi
    fi
    
    log "========================================"
    log "CalendarBridge Maintenance Starting"
    log "========================================"
    
    cleanup_logs
    echo "" >> "${HEALTH_LOG}"
    
    cleanup_outbox
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
