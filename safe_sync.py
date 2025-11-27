cat > ~/calendarBridge/safe_sync.py << 'SAFESYNC_EOF'
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CalendarBridge Safe Sync - Outlook → Google Calendar
Version 6.0.0 - Production Release

Features:
- State tracking for fast incremental syncs
- Conservative deletion (only events we created)
- Proper all-day event handling for iOS banner display
- Recurring event expansion within sync window
- Comprehensive error handling and logging
- Resilient to failures (continues processing, logs errors)
"""

import os
import sys
import json
import time
import random
import hashlib
import logging
from datetime import datetime, timedelta, timezone, date
from typing import Dict, Any, Optional, Tuple, List, Set
from pathlib import Path

# ---------------- Logging Setup ----------------

def setup_logging(log_dir: str) -> logging.Logger:
    """Configure logging to both file and console."""
    Path(log_dir).mkdir(parents=True, exist_ok=True)
    
    log_file = os.path.join(log_dir, f"sync_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    
    logger = logging.getLogger("calendarbridge")
    logger.setLevel(logging.DEBUG)
    
    # Clear any existing handlers
    logger.handlers = []
    
    # File handler - detailed
    fh = logging.FileHandler(log_file)
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(logging.Formatter('%(asctime)s [%(levelname)s] %(message)s'))
    
    # Console handler - info and above
    ch = logging.StreamHandler()
    ch.setLevel(logging.INFO)
    ch.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
    
    logger.addHandler(fh)
    logger.addHandler(ch)
    
    return logger

# ---------------- Configuration ----------------

def load_config(path: str) -> Dict[str, Any]:
    """Load configuration from JSON file."""
    if os.path.exists(path):
        with open(path, "r") as f:
            return json.load(f)
    return {}

ROOT = os.path.expanduser("~/calendarBridge")
CONFIG_PATH = os.path.join(ROOT, "calendar_config.json")
STATE_PATH = os.path.join(ROOT, "sync_state.json")
LOG_DIR = os.path.join(ROOT, "logs")

CONF = load_config(CONFIG_PATH)

API_DELAY = float(CONF.get("api_delay_seconds", 0.15))
TIMEZONE = CONF.get("timezone", "America/New_York")
CAL_ID = CONF.get("google_calendar_id", "primary")
DAYS_PAST = int(CONF.get("sync_days_past", 60))
DAYS_FUTURE = int(CONF.get("sync_days_future", 90))

# CRITICAL: Only delete events that have our marker
DELETE_ORPHANS = True
ORPHAN_MARKER = "calendarbridge"

# Environment overrides
QUOTA_USER = os.environ.get("CALBRIDGE_QUOTA_USER")
FORCE_FULL_SYNC = os.environ.get("CALBRIDGE_FORCE_FULL", "").lower() == "true"

# Initialize logger
log = setup_logging(LOG_DIR)

# Late imports
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from icalendar import Calendar
import recurring_ical_events
from dateutil.tz import gettz

# ---------------- State Management ----------------

class SyncState:
    """Manages sync state for incremental updates."""
    
    def __init__(self, path: str):
        self.path = path
        self.data = self._load()
    
    def _load(self) -> Dict[str, Any]:
        """Load state from disk."""
        if os.path.exists(self.path):
            try:
                with open(self.path, "r") as f:
                    return json.load(f)
            except Exception as e:
                log.warning(f"Failed to load state file, starting fresh: {e}")
        return {"events": {}, "google_ids": {}, "last_sync": None}
    
    def save(self):
        """Persist state to disk."""
        try:
            self.data["last_sync"] = datetime.now(timezone.utc).isoformat()
            with open(self.path, "w") as f:
                json.dump(self.data, f, indent=2)
        except Exception as e:
            log.error(f"Failed to save state: {e}")
    
    def get_event_hash(self, key: str) -> Optional[str]:
        """Get stored hash for an event."""
        return self.data.get("events", {}).get(key)
    
    def set_event_hash(self, key: str, content_hash: str, google_id: str):
        """Store hash and Google ID for an event."""
        if "events" not in self.data:
            self.data["events"] = {}
        if "google_ids" not in self.data:
            self.data["google_ids"] = {}
        self.data["events"][key] = content_hash
        self.data["google_ids"][key] = google_id
    
    def get_google_id(self, key: str) -> Optional[str]:
        """Get stored Google event ID for a local event key."""
        return self.data.get("google_ids", {}).get(key)
    
    def remove_event(self, key: str):
        """Remove an event from state."""
        self.data.get("events", {}).pop(key, None)
        self.data.get("google_ids", {}).pop(key, None)
    
    def get_all_keys(self) -> Set[str]:
        """Get all tracked event keys."""
        return set(self.data.get("events", {}).keys())

def compute_event_hash(ev: Dict[str, Any]) -> str:
    """Compute a hash of event content for change detection."""
    content = json.dumps({
        "summary": ev.get("summary", ""),
        "location": ev.get("location", ""),
        "description": ev.get("description", "")[:500],
        "start": str(ev.get("start", "")),
        "end": str(ev.get("end", "")),
        "allDay": ev.get("allDay", False)
    }, sort_keys=True)
    return hashlib.md5(content.encode()).hexdigest()

# ---------------- Google API Setup ----------------

def get_google_service():
    """Initialize Google Calendar API service."""
    creds_path = os.path.join(ROOT, "token.json")
    creds = Credentials.from_authorized_user_file(creds_path)
    return build("calendar", "v3", credentials=creds, cache_discovery=False)

# ---------------- Rate Limiting & Backoff ----------------

def sleep_with_jitter(base_seconds: float) -> None:
    """Sleep with +/-20% jitter."""
    if base_seconds <= 0:
        return
    jitter = base_seconds * random.uniform(-0.20, 0.20)
    time.sleep(max(0.0, base_seconds + jitter))

def is_rate_limit_error(e: HttpError) -> bool:
    """Check if error is a rate limit."""
    try:
        status = e.resp.status if hasattr(e, "resp") else None
        if status in (429,):
            return True
        if status == 403:
            content = getattr(e, "content", b"") or b""
            msg = content.decode("utf-8", errors="ignore")
            return "rateLimitExceeded" in msg or "userRateLimitExceeded" in msg
    except:
        pass
    return False

def request_with_backoff(service, req_builder, op_desc: str, max_attempts: int = 5) -> Any:
    """Execute Google API request with exponential backoff."""
    attempt = 0
    backoff = 1.0
    
    while True:
        attempt += 1
        try:
            sleep_with_jitter(API_DELAY)
            req = req_builder()
            if QUOTA_USER:
                req.uri += ("&" if "?" in req.uri else "?") + f"quotaUser={QUOTA_USER}"
            return req.execute()
        except HttpError as e:
            if is_rate_limit_error(e) and attempt < max_attempts:
                wait = backoff + random.uniform(0.0, 1.0)
                log.warning(f"Rate limited on {op_desc}, sleeping {wait:.1f}s (attempt {attempt})")
                time.sleep(wait)
                backoff = min(backoff * 2.0, 30.0)
                continue
            raise
        except Exception as e:
            if attempt < max_attempts:
                wait = backoff + random.uniform(0.0, 0.5)
                log.warning(f"Error on {op_desc}: {e}, retrying in {wait:.1f}s")
                time.sleep(wait)
                backoff = min(backoff * 2.0, 30.0)
                continue
            raise

# ---------------- ICS Parsing ----------------

def to_iso(dt) -> str:
    """Convert datetime/date to ISO 8601 string."""
    if isinstance(dt, date) and not isinstance(dt, datetime):
        return dt.isoformat()
    if hasattr(dt, "tzinfo") and dt.tzinfo is not None:
        return dt.isoformat()
    if isinstance(dt, datetime):
        return dt.replace(tzinfo=timezone.utc).isoformat()
    try:
        return datetime(dt.year, dt.month, dt.day, tzinfo=timezone.utc).isoformat()
    except:
        return str(dt)

def normalize_to_date(dt) -> date:
    """Convert to date object."""
    if isinstance(dt, datetime):
        return dt.date()
    if isinstance(dt, date):
        return dt
    return date(dt.year, dt.month, dt.day)

def is_all_day_event(comp) -> bool:
    """Detect if event is all-day using multiple methods."""
    # Method 1: Microsoft's flag
    ms_allday = comp.get("X-MICROSOFT-CDO-ALLDAYEVENT")
    if ms_allday and str(ms_allday).upper().strip() == "TRUE":
        return True
    
    dtstart = comp.get("DTSTART")
    if not dtstart:
        return False
    
    dt_start = dtstart.dt
    
    # Method 2: Pure date (VALUE=DATE)
    if isinstance(dt_start, date) and not isinstance(dt_start, datetime):
        return True
    
    # Method 3: Midnight-to-midnight
    if isinstance(dt_start, datetime):
        dtend = comp.get("DTEND")
        if dtend:
            dt_end = dtend.dt
            if isinstance(dt_end, datetime):
                start_midnight = dt_start.hour == 0 and dt_start.minute == 0 and dt_start.second == 0
                end_midnight = dt_end.hour == 0 and dt_end.minute == 0 and dt_end.second == 0
                if start_midnight and end_midnight and dt_end > dt_start:
                    delta = dt_end - dt_start
                    if delta.days >= 1 and delta.seconds == 0:
                        return True
    
    return False

def get_sync_window() -> Tuple[datetime, datetime]:
    """Get sync window as timezone-aware datetimes."""
    tz = gettz(TIMEZONE)
    now = datetime.now(tz)
    return now - timedelta(days=DAYS_PAST), now + timedelta(days=DAYS_FUTURE)

def load_local_events(outbox_dir: str) -> Dict[str, Dict[str, Any]]:
    """Parse ICS files and expand recurring events."""
    events: Dict[str, Dict[str, Any]] = {}
    window_start, window_end = get_sync_window()
    
    combined_path = os.path.join(outbox_dir, "outlook_full_export.ics")
    
    if not os.path.exists(combined_path):
        log.error(f"No ICS file found at {combined_path}")
        return events
    
    stats = {"all_day": 0, "timed": 0, "recurring": 0, "errors": 0}
    
    try:
        with open(combined_path, "rb") as f:
            raw_data = f.read()
        
        cal = Calendar.from_ical(raw_data)
        
        # Expand recurring events
        try:
            expanded = recurring_ical_events.of(cal).between(window_start, window_end)
        except Exception as e:
            log.warning(f"Recurrence expansion failed: {e}, falling back to simple parse")
            expanded = list(cal.walk("VEVENT"))
        
        for comp in expanded:
            try:
                if comp.name != "VEVENT":
                    continue
                
                uid = str(comp.get("UID") or "").strip()
                if not uid:
                    continue
                
                dtstart_prop = comp.get("DTSTART")
                dtend_prop = comp.get("DTEND")
                if not dtstart_prop:
                    continue
                
                dtstart = dtstart_prop.dt
                dtend = dtend_prop.dt if dtend_prop else None
                
                # Track recurring
                if comp.get("RRULE") or comp.get("RECURRENCE-ID"):
                    stats["recurring"] += 1
                
                all_day = is_all_day_event(comp)
                
                if all_day:
                    stats["all_day"] += 1
                    start_date = normalize_to_date(dtstart)
                    end_date = normalize_to_date(dtend) if dtend else start_date + timedelta(days=1)
                    if end_date <= start_date:
                        end_date = start_date + timedelta(days=1)
                    
                    key = f"{uid}|{start_date.isoformat()}"
                    events[key] = {
                        "uid": uid,
                        "summary": (comp.get("SUMMARY") or "").strip(),
                        "location": (comp.get("LOCATION") or "").strip(),
                        "description": (comp.get("DESCRIPTION") or "").strip(),
                        "start": start_date,
                        "end": end_date,
                        "allDay": True
                    }
                else:
                    stats["timed"] += 1
                    if dtend is None:
                        dtend = dtstart + timedelta(hours=1) if isinstance(dtstart, datetime) else dtstart + timedelta(days=1)
                    
                    key = f"{uid}|{to_iso(dtstart)}"
                    events[key] = {
                        "uid": uid,
                        "summary": (comp.get("SUMMARY") or "").strip(),
                        "location": (comp.get("LOCATION") or "").strip(),
                        "description": (comp.get("DESCRIPTION") or "").strip(),
                        "start": dtstart,
                        "end": dtend,
                        "allDay": False
                    }
            except Exception as e:
                stats["errors"] += 1
                log.debug(f"Error parsing event: {e}")
                continue
    
    except Exception as e:
        log.error(f"Failed to parse ICS file: {e}")
        return events
    
    log.info(f"Parsed {len(events)} events: {stats['all_day']} all-day, {stats['timed']} timed, {stats['recurring']} recurring instances, {stats['errors']} errors")
    return events

# ---------------- Google Event Operations ----------------

def build_event_body(ev: Dict[str, Any], tz: str) -> Dict[str, Any]:
    """Build Google Calendar event body."""
    body = {
        "summary": ev["summary"] or "(No title)",
        "location": ev["location"] or None,
        "description": ev["description"] or None,
        "extendedProperties": {
            "private": {
                "icalUID": ev["uid"],
                "source": ORPHAN_MARKER
            }
        }
    }
    
    if ev["allDay"]:
        start_date = ev["start"]
        end_date = ev["end"]
        if isinstance(start_date, datetime):
            start_date = start_date.date()
        if isinstance(end_date, datetime):
            end_date = end_date.date()
        
        body["start"] = {"date": start_date.isoformat()}
        body["end"] = {"date": end_date.isoformat()}
        body["transparency"] = "transparent"
    else:
        body["start"] = {"dateTime": to_iso(ev["start"]), "timeZone": tz}
        body["end"] = {"dateTime": to_iso(ev["end"]), "timeZone": tz}
    
    return body

def get_time_window_iso() -> Tuple[str, str]:
    """Get ISO timestamps for API calls."""
    now = datetime.now(timezone.utc)
    time_min = now - timedelta(days=DAYS_PAST)
    time_max = now + timedelta(days=DAYS_FUTURE)
    return time_min.isoformat(), time_max.isoformat()

def fetch_google_events(service, calendar_id: str, time_min: str, time_max: str) -> Dict[str, Dict[str, Any]]:
    """Fetch Google events and index by iCalUID+start."""
    events_by_key: Dict[str, Dict[str, Any]] = {}
    page_token = None
    
    while True:
        try:
            resp = request_with_backoff(
                service,
                lambda: service.events().list(
                    calendarId=calendar_id,
                    timeMin=time_min,
                    timeMax=time_max,
                    singleEvents=True,
                    showDeleted=False,
                    maxResults=2500,
                    pageToken=page_token,
                    orderBy="startTime"
                ),
                "events.list"
            )
        except Exception as e:
            log.error(f"Failed to fetch Google events: {e}")
            break
        
        for item in resp.get("items", []):
            ical_uid = item.get("iCalUID")
            if not ical_uid:
                ext = item.get("extendedProperties", {}).get("private", {}) or {}
                ical_uid = ext.get("icalUID")
            
            if ical_uid:
                start = item.get("start", {})
                start_str = start.get("date") or start.get("dateTime", "")
                if start_str:
                    if "T" in start_str:
                        parts = start_str.split("T")
                        time_part = parts[1][:8] if len(parts[1]) >= 8 else parts[1].split("+")[0].split("-")[0]
                        start_str = f"{parts[0]}T{time_part}"
                    key = f"{ical_uid}|{start_str}"
                    events_by_key[key] = item
        
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    
    log.info(f"Fetched {len(events_by_key)} events from Google Calendar")
    return events_by_key

def is_our_event(g_event: Dict[str, Any]) -> bool:
    """Check if we created this event (has our marker)."""
    ext = g_event.get("extendedProperties", {}).get("private", {}) or {}
    return ext.get("source") == ORPHAN_MARKER

def upsert_event(service, calendar_id: str, ev: Dict[str, Any], body: Dict[str, Any], 
                 existing: Optional[Dict[str, Any]], state: SyncState, key: str) -> Tuple[str, Optional[str]]:
    """Create or update an event. Returns (action, google_id)."""
    if existing:
        existing_id = existing.get("id")
        
        current_summary = existing.get("summary", "")
        current_location = existing.get("location", "")
        current_start = existing.get("start", {})
        current_end = existing.get("end", {})
        current_transparency = existing.get("transparency", "opaque")
        
        needs_update = False
        
        if body.get("summary") != current_summary:
            needs_update = True
        elif body.get("location") != current_location:
            needs_update = True
        elif body.get("transparency") != current_transparency:
            needs_update = True
        elif body.get("start") != current_start:
            needs_update = True
        elif body.get("end") != current_end:
            needs_update = True
        
        if not needs_update:
            return "skipped", existing_id
        
        try:
            resp = request_with_backoff(
                service,
                lambda: service.events().patch(
                    calendarId=calendar_id,
                    eventId=existing_id,
                    body=body,
                    sendUpdates="none"
                ),
                f"patch {ev['summary'][:20] if ev.get('summary') else 'event'}"
            )
            return "updated", resp.get("id")
        except HttpError as e:
            if e.resp.status == 404:
                pass
            else:
                log.error(f"Patch failed for {key}: {e.resp.status}")
                return "failed", None
        except Exception as e:
            log.error(f"Patch error for {key}: {e}")
            return "failed", None
    
    try:
        resp = request_with_backoff(
            service,
            lambda: service.events().insert(
                calendarId=calendar_id,
                body=body
            ),
            f"insert {ev['summary'][:20] if ev.get('summary') else 'event'}"
        )
        return "created", resp.get("id")
    except HttpError as e:
        if e.resp.status == 409:
            return "skipped", None
        log.error(f"Insert failed for {key}: {e.resp.status}")
        return "failed", None
    except Exception as e:
        log.error(f"Insert error for {key}: {e}")
        return "failed", None

# ---------------- Main Sync ----------------

def main():
    """Main sync entry point."""
    start_time = time.time()
    
    log.info("=" * 60)
    log.info(f"CalendarBridge Safe Sync v6.0.0")
    log.info(f"Calendar: {CAL_ID}")
    log.info(f"Window: -{DAYS_PAST} to +{DAYS_FUTURE} days")
    log.info(f"Timezone: {TIMEZONE}")
    log.info("=" * 60)
    
    state = SyncState(STATE_PATH)
    service = get_google_service()
    
    outbox = os.path.join(ROOT, "outbox")
    local_events = load_local_events(outbox)
    
    if not local_events:
        log.error("No events parsed from ICS - aborting to prevent data loss")
        sys.exit(1)
    
    time_min, time_max = get_time_window_iso()
    google_events = fetch_google_events(service, CAL_ID, time_min, time_max)
    
    stats = {"created": 0, "updated": 0, "skipped": 0, "deleted": 0, "failed": 0}
    
    local_keys = set(local_events.keys())
    processed_google_ids = set()
    
    for key, ev in local_events.items():
        try:
            content_hash = compute_event_hash(ev)
            stored_hash = state.get_event_hash(key)
            stored_google_id = state.get_google_id(key)
            
            existing = None
            
            if key in google_events:
                existing = google_events[key]
            elif stored_google_id:
                for g_key, g_ev in google_events.items():
                    if g_ev.get("id") == stored_google_id:
                        existing = g_ev
                        break
            
            if existing and stored_hash == content_hash and not FORCE_FULL_SYNC:
                stats["skipped"] += 1
                if existing.get("id"):
                    processed_google_ids.add(existing["id"])
                continue
            
            body = build_event_body(ev, TIMEZONE)
            action, google_id = upsert_event(service, CAL_ID, ev, body, existing, state, key)
            
            stats[action] += 1
            
            if google_id:
                state.set_event_hash(key, content_hash, google_id)
                processed_google_ids.add(google_id)
            
            if existing and existing.get("id"):
                processed_google_ids.add(existing["id"])
                
        except Exception as e:
            log.error(f"Error processing event {key}: {e}")
            stats["failed"] += 1
            continue
    
    if DELETE_ORPHANS:
        state_keys = state.get_all_keys()
        orphan_keys = state_keys - local_keys
        
        for key in orphan_keys:
            google_id = state.get_google_id(key)
            if not google_id:
                state.remove_event(key)
                continue
            
            found_event = None
            for g_key, g_ev in google_events.items():
                if g_ev.get("id") == google_id:
                    found_event = g_ev
                    break
            
            if found_event and is_our_event(found_event):
                try:
                    request_with_backoff(
                        service,
                        lambda gid=google_id: service.events().delete(
                            calendarId=CAL_ID,
                            eventId=gid,
                            sendUpdates="none"
                        ),
                        f"delete orphan"
                    )
                    stats["deleted"] += 1
                    log.debug(f"Deleted orphan: {found_event.get('summary', 'unknown')}")
                except HttpError as e:
                    if e.resp.status == 410:
                        stats["deleted"] += 1
                    else:
                        log.warning(f"Delete failed for {google_id}: {e.resp.status}")
                except Exception as e:
                    log.warning(f"Delete error for {google_id}: {e}")
            
            state.remove_event(key)
    
    state.save()
    
    elapsed = time.time() - start_time
    log.info("=" * 60)
    log.info(f"SYNC COMPLETE in {elapsed:.1f}s")
    log.info(f"Created: {stats['created']}, Updated: {stats['updated']}, Skipped: {stats['skipped']}, Deleted: {stats['deleted']}, Failed: {stats['failed']}")
    log.info("=" * 60)
    
    if stats["failed"] > 0:
        sys.exit(1)

if __name__ == "__main__":
    main()
SAFESYNC_EOF
