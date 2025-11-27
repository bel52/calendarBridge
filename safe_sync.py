#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CalendarBridge Safe Sync - Outlook → Google Calendar
Version 6.1.0 - Production Release

Features:
- Multi-VCALENDAR block support for Outlook exports
- Deterministic key per event instance: UID|normalized start
- Fixed timezone key matching to prevent duplicates
- Idempotent: no duplicate creation for the same UID/start
- Conservative deletion (only events we created)
- All-day vs timed event handling for proper iOS display
- Recurring expansion within sync window
"""

import os
import sys
import json
import time
import logging
import hashlib
from dataclasses import dataclass
from typing import Dict, Any, Optional, Tuple, Union, List
from datetime import datetime, date, timedelta, timezone

try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo

from icalendar import Calendar
from recurring_ical_events import of as recurring_of

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# ============================================================================
# Configuration
# ============================================================================

ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(ROOT, "calendar_config.json")
STATE_PATH = os.path.join(ROOT, "sync_state.json")
TOKEN_PATH = os.path.join(ROOT, "token.json")
CREDENTIALS_PATH = os.path.join(ROOT, "credentials.json")
LOG_DIR = os.path.join(ROOT, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

LOG_PATH = os.path.join(LOG_DIR, "safe_sync.log")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH),
        logging.StreamHandler(sys.stdout),
    ],
)

log = logging.getLogger("calendarbridge")

ORPHAN_MARKER = "CalendarBridge"
VERSION = "6.1.0"

# ============================================================================
# Data Models
# ============================================================================

@dataclass
class LocalEvent:
    uid: str
    key: str
    summary: str
    location: str
    description: str
    start: Union[datetime, date]
    end: Union[datetime, date]
    all_day: bool


# ============================================================================
# Configuration & Authentication
# ============================================================================

def load_config() -> Dict[str, Any]:
    if not os.path.exists(CONFIG_PATH):
        raise SystemExit(f"Missing config file: {CONFIG_PATH}")
    with open(CONFIG_PATH, "r") as f:
        cfg = json.load(f)
    required = ["google_calendar_id", "timezone", "sync_days_past", "sync_days_future"]
    missing = [k for k in required if k not in cfg]
    if missing:
        raise SystemExit(f"Missing config keys: {missing}")
    return cfg


def get_timezone(tz_name: str) -> ZoneInfo:
    try:
        return ZoneInfo(tz_name)
    except Exception as e:
        raise SystemExit(f"Invalid timezone '{tz_name}': {e}")


def get_google_service(scopes: Optional[List[str]] = None):
    if scopes is None:
        scopes = ["https://www.googleapis.com/auth/calendar"]
    creds = None
    if os.path.exists(TOKEN_PATH):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_PATH, scopes)
        except Exception as e:
            log.warning(f"Failed to load token.json, re-authenticating: {e}")
            creds = None

    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                log.warning(f"Token refresh failed, re-authenticating: {e}")
                creds = None

        if not creds:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, scopes)
            creds = flow.run_local_server(port=0)

        with open(TOKEN_PATH, "w") as token:
            token.write(creds.to_json())

    return build("calendar", "v3", credentials=creds, cache_discovery=False)


# ============================================================================
# State Management
# ============================================================================

class SyncState:
    def __init__(self, path: str):
        self.path = path
        self.data = {"events": {}, "google_ids": {}, "last_sync": None}
        self._load()

    def _load(self):
        if os.path.exists(self.path):
            try:
                with open(self.path, "r") as f:
                    self.data = json.load(f)
            except Exception as e:
                log.warning(f"Failed to load state file, starting fresh: {e}")
                self.data = {"events": {}, "google_ids": {}, "last_sync": None}

    def save(self):
        try:
            self.data["last_sync"] = datetime.now(timezone.utc).isoformat()
            with open(self.path, "w") as f:
                json.dump(self.data, f, indent=2)
        except Exception as e:
            log.error(f"Failed to save state: {e}")

    def get_hash(self, key: str) -> Optional[str]:
        return self.data.get("events", {}).get(key)

    def set_hash(self, key: str, content_hash: str, google_id: str):
        self.data.setdefault("events", {})[key] = content_hash
        self.data.setdefault("google_ids", {})[key] = google_id

    def get_google_id(self, key: str) -> Optional[str]:
        return self.data.get("google_ids", {}).get(key)

    def remove(self, key: str):
        self.data.get("events", {}).pop(key, None)
        self.data.get("google_ids", {}).pop(key, None)


# ============================================================================
# Time Utilities
# ============================================================================

def get_sync_window(cfg: Dict[str, Any], tz: ZoneInfo) -> Tuple[datetime, datetime]:
    now = datetime.now(tz)
    start = now - timedelta(days=int(cfg["sync_days_past"]))
    end = now + timedelta(days=int(cfg["sync_days_future"]))
    return start.replace(tzinfo=None), end.replace(tzinfo=None)


def normalize_to_date(dt_or_date: Union[datetime, date], tz: ZoneInfo) -> date:
    if isinstance(dt_or_date, date) and not isinstance(dt_or_date, datetime):
        return dt_or_date
    if isinstance(dt_or_date, datetime):
        if dt_or_date.tzinfo is None:
            dt_or_date = dt_or_date.replace(tzinfo=tz)
        return dt_or_date.astimezone(tz).date()
    raise TypeError(f"Unsupported type: {type(dt_or_date)}")


def normalize_to_datetime(dt: Union[datetime, date], tz: ZoneInfo) -> datetime:
    if isinstance(dt, datetime):
        if dt.tzinfo is None:
            return dt.replace(tzinfo=tz)
        return dt.astimezone(tz)
    if isinstance(dt, date):
        return datetime(dt.year, dt.month, dt.day, tzinfo=tz)
    raise TypeError(f"Unsupported type: {type(dt)}")


def is_all_day_event(comp) -> bool:
    """Detect all-day events using multiple heuristics"""
    ms_allday = comp.get("X-MICROSOFT-CDO-ALLDAYEVENT")
    if ms_allday and str(ms_allday).strip().upper() == "TRUE":
        return True

    dtstart_prop = comp.get("DTSTART")
    if not dtstart_prop:
        return False

    dtstart = dtstart_prop.dt
    if isinstance(dtstart, date) and not isinstance(dtstart, datetime):
        return True

    dtend_prop = comp.get("DTEND")
    if isinstance(dtstart, datetime) and dtend_prop is not None:
        dtend = dtend_prop.dt
        if isinstance(dtend, datetime):
            start_midnight = dtstart.hour == 0 and dtstart.minute == 0 and dtstart.second == 0
            end_midnight = dtend.hour == 0 and dtend.minute == 0 and dtend.second == 0
            if start_midnight and end_midnight and dtend > dtstart:
                delta = dtend - dtstart
                if delta.days >= 1 and delta.seconds == 0:
                    return True

    return False


def compute_event_hash(ev: LocalEvent, tz: ZoneInfo) -> str:
    """Generate content hash for change detection"""
    if ev.all_day:
        start_repr = normalize_to_date(ev.start, tz).isoformat()
        end_repr = normalize_to_date(ev.end, tz).isoformat()
    else:
        start_repr = normalize_to_datetime(ev.start, tz).isoformat(timespec="seconds")
        end_repr = normalize_to_datetime(ev.end, tz).isoformat(timespec="seconds")

    payload = {
        "uid": ev.uid,
        "summary": ev.summary,
        "location": ev.location,
        "description": ev.description,
        "all_day": ev.all_day,
        "start": start_repr,
        "end": end_repr,
    }
    data = json.dumps(payload, sort_keys=True, ensure_ascii=False)
    return hashlib.sha256(data.encode("utf-8")).hexdigest()


def to_iso(dt: datetime) -> str:
    """Convert datetime to ISO format in UTC"""
    if dt.tzinfo is None:
        dt = dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc).isoformat()


# ============================================================================
# ICS Parsing with Multi-VCALENDAR Support
# ============================================================================

def load_local_events(cfg: Dict[str, Any], tz: ZoneInfo) -> Dict[str, LocalEvent]:
    """
    Parse Outlook ICS export which may contain multiple VCALENDAR blocks.
    Returns dict keyed by UID|normalized_start_time
    """
    outbox_dir = os.path.join(ROOT, "outbox")
    ics_path = os.path.join(outbox_dir, "outlook_full_export.ics")

    if not os.path.exists(ics_path):
        log.error(f"No ICS found at {ics_path}")
        return {}

    window_start, window_end = get_sync_window(cfg, tz)
    log.info(f"Sync window (local time, naive): {window_start} → {window_end}")

    events: Dict[str, LocalEvent] = {}
    stats = {"all_day": 0, "timed": 0, "recurring": 0, "errors": 0}

    # Read and split by VCALENDAR blocks
    with open(ics_path, "rb") as f:
        raw_data = f.read()

    content = raw_data.decode("utf-8", errors="ignore")
    
    # Split into individual VCALENDAR blocks
    blocks = content.split("BEGIN:VCALENDAR")
    vcal_count = 0
    
    for block in blocks:
        if not block.strip():
            continue
        
        # Reconstruct valid VCALENDAR
        ics_block = "BEGIN:VCALENDAR" + block
        if not ics_block.rstrip().endswith("END:VCALENDAR"):
            ics_block = ics_block.rstrip() + "\nEND:VCALENDAR"
        
        try:
            cal = Calendar.from_ical(ics_block.encode("utf-8"))
            vcal_count += 1
        except Exception as e:
            log.debug(f"Skipping invalid VCALENDAR block: {e}")
            continue

        # Expand recurring events
        try:
            expanded = recurring_of(cal).between(window_start, window_end)
        except Exception as e:
            log.warning(f"Failed to expand recurrences in block: {e}")
            expanded = list(cal.walk("VEVENT"))

        for comp in expanded:
            if comp.name != "VEVENT":
                continue
            
            try:
                uid = str(comp.get("UID") or "").strip()
                if not uid:
                    continue

                dtstart_prop = comp.get("DTSTART")
                dtend_prop = comp.get("DTEND")
                if not dtstart_prop:
                    continue

                dtstart = dtstart_prop.dt
                dtend = dtend_prop.dt if dtend_prop is not None else None

                if comp.get("RRULE") or comp.get("RECURRENCE-ID"):
                    stats["recurring"] += 1

                all_day = is_all_day_event(comp)

                if all_day:
                    stats["all_day"] += 1
                    start_date = normalize_to_date(dtstart, tz)
                    end_date = normalize_to_date(dtend, tz) if dtend else start_date + timedelta(days=1)
                    if end_date <= start_date:
                        end_date = start_date + timedelta(days=1)

                    start_key = start_date.isoformat()
                    key = f"{uid}|{start_key}"

                    events[key] = LocalEvent(
                        uid=uid,
                        key=key,
                        summary=str(comp.get("SUMMARY") or "").strip(),
                        location=str(comp.get("LOCATION") or "").strip(),
                        description=str(comp.get("DESCRIPTION") or "").strip(),
                        start=start_date,
                        end=end_date,
                        all_day=True,
                    )
                else:
                    stats["timed"] += 1
                    start_dt = normalize_to_datetime(dtstart, tz)
                    if dtend is not None:
                        end_dt = normalize_to_datetime(dtend, tz)
                    else:
                        end_dt = start_dt + timedelta(hours=1)

                    # CRITICAL FIX: Strip timezone to match Google key format
                    start_iso = start_dt.isoformat(timespec="seconds")
                    if "T" in start_iso:
                        date_part, time_part = start_iso.split("T", 1)
                        time_part = time_part.split("+")[0].split("-")[0]
                        start_key = f"{date_part}T{time_part[:8]}"
                    else:
                        start_key = start_iso
                    key = f"{uid}|{start_key}"

                    events[key] = LocalEvent(
                        uid=uid,
                        key=key,
                        summary=str(comp.get("SUMMARY") or "").strip(),
                        location=str(comp.get("LOCATION") or "").strip(),
                        description=str(comp.get("DESCRIPTION") or "").strip(),
                        start=start_dt,
                        end=end_dt,
                        all_day=False,
                    )

            except Exception as e:
                stats["errors"] += 1
                log.debug(f"Error parsing event: {e}")
                continue

    log.info(f"Parsed {vcal_count} VCALENDAR blocks")
    log.info(
        f"Parsed {len(events)} events: {stats['all_day']} all-day, "
        f"{stats['timed']} timed, {stats['recurring']} recurring instances, "
        f"{stats['errors']} errors"
    )
    return events


# ============================================================================
# Google Calendar Operations
# ============================================================================

def build_event_body(ev: LocalEvent, tz_name: str) -> Dict[str, Any]:
    """Build Google Calendar API event body"""
    body: Dict[str, Any] = {
        "summary": ev.summary or "(No title)",
        "location": ev.location or None,
        "description": ev.description or None,
        "extendedProperties": {
            "private": {
                "icalUID": ev.uid,
                "source": ORPHAN_MARKER,
            }
        },
    }

    if ev.all_day:
        start_date = ev.start if isinstance(ev.start, date) else ev.start.date()
        end_date = ev.end if isinstance(ev.end, date) else ev.end.date()
        body["start"] = {"date": start_date.isoformat()}
        body["end"] = {"date": end_date.isoformat()}
        body["transparency"] = "transparent"
    else:
        start_dt = ev.start if isinstance(ev.start, datetime) else normalize_to_datetime(ev.start, ZoneInfo(tz_name))
        end_dt = ev.end if isinstance(ev.end, datetime) else normalize_to_datetime(ev.end, ZoneInfo(tz_name))
        body["start"] = {"dateTime": to_iso(start_dt), "timeZone": tz_name}
        body["end"] = {"dateTime": to_iso(end_dt), "timeZone": tz_name}

    return body


def get_time_window_iso(cfg: Dict[str, Any]) -> Tuple[str, str]:
    """Get sync window as ISO strings in UTC"""
    now = datetime.now(timezone.utc)
    time_min = now - timedelta(days=int(cfg["sync_days_past"]))
    time_max = now + timedelta(days=int(cfg["sync_days_future"]))
    return time_min.isoformat(), time_max.isoformat()


def _normalize_start_for_key(start: Dict[str, Any]) -> Optional[str]:
    """
    Normalize Google event start time to match local key format.
    Strips timezone and subseconds for consistent comparison.
    """
    if not start:
        return None
    if "date" in start:
        return start["date"]
    dt_str = start.get("dateTime")
    if not dt_str:
        return None
    if "T" not in dt_str:
        return dt_str
    date_part, time_part = dt_str.split("T", 1)
    time_part = time_part.split("+")[0].split("-")[0]
    time_part = time_part[:8]
    return f"{date_part}T{time_part}"


def fetch_google_events(service, calendar_id: str, cfg: Dict[str, Any]) -> Dict[str, Dict[str, Any]]:
    """Fetch all Google events in sync window, keyed by UID|start"""
    time_min, time_max = get_time_window_iso(cfg)

    events_by_key: Dict[str, Dict[str, Any]] = {}
    page_token = None
    total = 0

    while True:
        try:
            resp = (
                service.events()
                .list(
                    calendarId=calendar_id,
                    timeMin=time_min,
                    timeMax=time_max,
                    singleEvents=True,
                    showDeleted=False,
                    maxResults=2500,
                    pageToken=page_token,
                    orderBy="startTime",
                )
                .execute()
            )
        except HttpError as e:
            log.error(f"Failed to fetch Google events: {e}")
            break

        items = resp.get("items", [])
        total += len(items)

        for item in items:
            ext = (item.get("extendedProperties") or {}).get("private") or {}
            ical_uid = ext.get("icalUID") or item.get("iCalUID")
            if not ical_uid:
                continue

            start_key = _normalize_start_for_key(item.get("start") or {})
            if not start_key:
                continue

            key = f"{ical_uid}|{start_key}"
            events_by_key[key] = item

        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    log.info(f"Fetched {total} events from Google Calendar")
    return events_by_key


def is_our_event(item: Dict[str, Any]) -> bool:
    """Check if event was created by CalendarBridge"""
    ext = (item.get("extendedProperties") or {}).get("private") or {}
    return ext.get("source") == ORPHAN_MARKER


def safe_api_call(func, label: str, delay: float):
    """Execute API call with error handling and rate limiting"""
    try:
        result = func.execute()
        if delay > 0:
            time.sleep(delay)
        return result
    except HttpError as e:
        log.error(f"{label} failed: {e}")
        raise


# ============================================================================
# Sync Operations
# ============================================================================

def upsert_event(
    service,
    calendar_id: str,
    ev: LocalEvent,
    body: Dict[str, Any],
    existing: Optional[Dict[str, Any]],
    state: SyncState,
    content_hash: str,
    api_delay: float,
) -> Tuple[str, str]:
    """Create or update event, return (action, google_id)"""
    if existing:
        gid = existing.get("id")
        
        # EFFICIENCY: Check content hash first (no API call needed)
        stored_hash = state.get_hash(ev.key)
        if stored_hash == content_hash:
            # Content unchanged since last sync - skip API call entirely
            return "skipped", gid
        
        # Content changed - need to update
        log.info(f"Updating event {gid} ({ev.summary[:40]})")
        updated = safe_api_call(
            service.events().patch(calendarId=calendar_id, eventId=gid, body=body),
            "events.patch",
            api_delay,
        )
        state.set_hash(ev.key, content_hash, gid)
        return "updated", updated["id"]

    # Create new event
    log.info(f"Creating event: {ev.summary[:40]}")
    created = safe_api_call(
        service.events().insert(calendarId=calendar_id, body=body),
        "events.insert",
        api_delay,
    )
    gid = created["id"]
    state.set_hash(ev.key, content_hash, gid)
    return "created", gid


def delete_event(service, calendar_id: str, gid: str, api_delay: float):
    """Delete event from Google Calendar"""
    log.info(f"Deleting event {gid}")
    safe_api_call(
        service.events().delete(calendarId=calendar_id, eventId=gid),
        "events.delete",
        api_delay,
    )


# ============================================================================
# Main Sync Logic
# ============================================================================

def main():
    cfg = load_config()
    tz = get_timezone(cfg["timezone"])
    cal_id = cfg["google_calendar_id"]
    api_delay = float(cfg.get("api_delay_seconds", 0.05))

    log.info("=" * 60)
    log.info(f"CalendarBridge Safe Sync v{VERSION}")
    log.info(f"Calendar ID: {cal_id}")
    log.info(f"Timezone: {cfg['timezone']}")
    log.info("=" * 60)

    state = SyncState(STATE_PATH)
    service = get_google_service()

    # Parse local events
    local_events = load_local_events(cfg, tz)
    if not local_events:
        log.error("No local events parsed; aborting to avoid destructive sync")
        raise SystemExit(1)

    # Fetch Google events
    google_events = fetch_google_events(service, cal_id, cfg)

    stats = {"created": 0, "updated": 0, "skipped": 0, "deleted": 0, "failed": 0}
    processed_google_ids = set()
    start_time = time.time()

    # Upsert local → Google
    for key, ev in local_events.items():
        try:
            body = build_event_body(ev, cfg["timezone"])
            content_hash = compute_event_hash(ev, tz)

            existing = google_events.get(key)
            action, gid = upsert_event(
                service,
                cal_id,
                ev,
                body,
                existing,
                state,
                content_hash,
                api_delay,
            )
            stats[action] += 1
            processed_google_ids.add(gid)
            
            # Log progress every 50 events for long syncs
            if (stats["created"] + stats["updated"] + stats["skipped"]) % 50 == 0:
                log.debug(
                    f"Progress: {stats['created']} created, {stats['updated']} updated, "
                    f"{stats['skipped']} skipped"
                )
        except Exception as e:
            stats["failed"] += 1
            log.error(f"Failed to sync event {key}: {e}", exc_info=True)

    # Delete orphaned events (in Google but not in local)
    local_keys = set(local_events.keys())
    for key, item in google_events.items():
        if key in local_keys:
            continue
        if not is_our_event(item):
            continue
        gid = item.get("id")
        if not gid or gid in processed_google_ids:
            continue
        try:
            delete_event(service, cal_id, gid, api_delay)
            stats["deleted"] += 1
            state.remove(key)
        except Exception as e:
            stats["failed"] += 1
            log.error(f"Failed to delete event {gid}: {e}", exc_info=True)

    state.save()

    elapsed = time.time() - start_time
    log.info("=" * 60)
    log.info(f"SYNC COMPLETE in {elapsed:.1f}s")
    log.info(
        f"Created: {stats['created']}, Updated: {stats['updated']}, "
        f"Skipped: {stats['skipped']}, Deleted: {stats['deleted']}, "
        f"Failed: {stats['failed']}"
    )
    log.info("=" * 60)

    if stats["failed"] > 0:
        raise SystemExit(1)


if __name__ == "__main__":
    main()
