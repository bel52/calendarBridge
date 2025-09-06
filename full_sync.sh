#!/usr/bin/env bash
set -euo pipefail

CB="$HOME/calendarBridge"
LOGDIR="$CB/logs"
OUTBOX="$CB/outbox"
mkdir -p "$LOGDIR" "$OUTBOX"

ts() { date +"%Y-%m-%d %H:%M:%S"; }
log() { echo "$(ts) $*"; }

retry() {
  # retry <max_tries> <sleep_seconds> <command...>
  local tries="$1"; shift
  local sleep_s="$1"; shift
  local n=1
  while true; do
    if "$@"; then
      return 0
    fi
    if (( n >= tries )); then
      return 1
    fi
    log "WARN: Attempt $n failed. Retrying in ${sleep_s}s: $*"
    sleep "${sleep_s}"
    ((n++))
  done
}

check_connectivity() {
  /usr/bin/curl -I -sS --max-time 5 https://oauth2.googleapis.com/generate_204 >/dev/null || \
  /usr/bin/curl -I -sS --max-time 5 https://www.google.com/generate_204 >/dev/null
}

run_export() {
  osascript "$CB/exportEvents.scpt" "Calendar" 2
}

run_sync() {
  if [ -x "$CB/.venv/bin/python3" ]; then
    "$CB/.venv/bin/python3" "$CB/calendar_sync.py"
  else
    python3 "$CB/calendar_sync.py"
  fi
}

{
  log "--- CalendarBridge full_sync.sh start ---"

  log "Connectivity preflight to Google..."
  if ! retry 3 10 check_connectivity; then
    log "WARN: Connectivity check failed after retries; proceeding anyway."
  else
    log "Connectivity OK."
  fi

  log "Exporting from Outlook (with retries)..."
  if ! retry 3 10 run_export; then
    log "ERROR: AppleScript export failed after retries; aborting this run."
    log "--- full_sync.sh finished (failure) ---"
    exit 1
  fi
  log "Export complete. Latest outbox listing:"
  ls -lh "$OUTBOX" | tail -n +1

  log "Running calendar_sync.py (with retries)..."
  if ! retry 3 10 run_sync; then
    log "ERROR: calendar_sync.py failed after retries."
    log "--- full_sync.sh finished (failure) ---"
    exit 1
  fi

  log "--- full_sync.sh finished (success) ---"
} >> "$LOGDIR/launchd_out.log" 2>> "$LOGDIR/launchd_err.log"
