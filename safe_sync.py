#!/usr/bin/env python3
import os
import sys
import json
import hashlib
import random
import time
import logging
from datetime import datetime, timedelta, timezone
from typing import Dict, Any, List
from dateutil.tz import gettz
import recurring_ical_events as rie
from icalendar import Calendar
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow

# ----------------------------------------------------------------------
# Logging configuration
# ----------------------------------------------------------------------
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("safe_sync")

# ----------------------------------------------------------------------
# Helper functions
# ----------------------------------------------------------------------
def safe_event_id(key: str) -> str:
    """Generate a deterministic event ID from a composite key."""
    digest = hashlib.sha1(key.encode("utf-8")).hexdigest()[:32]
    return f"e{digest}"

def gexec(request_callable, verb_hint: str = "", uid_hint: str = "", delay: float = 0.05):
    """Execute a Google API call with basic rate‑limit handling."""
    attempt = 0
    while True:
        time.sleep(delay)
        try:
            return request_callable.execute()
        except HttpError as e:
            attempt += 1
            status = getattr(e.resp, "status", None)
            if status in (403, 429) and attempt <= 5:
                sleep_s = min(2.0 * attempt, 6.0) + random.uniform(0.1, 0.5)
                log.info(f"[rate] {verb_hint} uid={uid_hint} got {status}; sleeping {sleep_s:.1f}s (attempt {attempt}/5)")
                time.sleep(sleep_s)
                continue
            raise

def parse_ics(path: str, tzname: str, start_window: datetime, end_window: datetime) -> List[Dict[str, Any]]:
    """Parse a cleaned .ics file and return occurrences in the desired window."""
    tz = gettz(tzname)
    instances: List[Dict[str, Any]] = []
    with open(path, "rb") as f:
        cal = Calendar.from_ical(f.read())
    try:
        occs = rie.of(cal).between(start_window, end_window)
    except Exception as e:
        log.info(f"[WARN] Failed to expand recurrences in {path}: {e}")
        return instances

    for comp in occs:
        if comp.name != "VEVENT":
            continue
        uid = str(comp.get("UID") or "").strip()
        if not uid:
            continue
        summary = str(comp.get("SUMMARY") or "").strip()
        description = str(comp.get("DESCRIPTION") or "").strip()
        location = str(comp.get("LOCATION") or "").strip()
        dtstart = comp.get("DTSTART").dt if comp.get("DTSTART") else None
        dtend = comp.get("DTEND").dt if comp.get("DTEND") else None
        all_day = False
        if isinstance(dtstart, datetime):
            dtstart = dtstart.astimezone(gettz(tzname))
        else:
            all_day = True
            dtstart = tz.localize(datetime(dtstart.year, dtstart.month, dtstart.day))
        if isinstance(dtend, datetime):
            dtend = dtend.astimezone(gettz(tzname))
        else:
            if dtend is not None:
                dtend = tz.localize(datetime(dtend.year, dtend.month, dtend.day))
            else:
                dtend = dtstart + timedelta(hours=1)
        instances.append({
            "uid": uid,
            "summary": summary,
            "description": description,
            "location": location,
            "start": dtstart,
            "end": dtend,
            "all_day": all_day,
        })
    return instances

def load_config() -> Dict[str, Any]:
    """Always load calendar_config.json relative to this script."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    cfg_path = os.path.join(base_dir, "calendar_config.json")
    with open(cfg_path) as f:
        return json.load(f)

def get_service():
    """Obtain or refresh Google credentials, prompting once if token is missing."""
    scopes = ["https://www.googleapis.com/auth/calendar"]
    creds = None
    token_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "token.json")
    creds_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, scopes)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                pass
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file(creds_path, scopes)
            creds = flow.run_local_server(port=0)
        with open(token_path, "w") as token_file:
            token_file.write(creds.to_json())
    return build("calendar", "v3", credentials=creds, cache_discovery=False)

def load_state(path: str) -> Dict[str, str]:
    """Load sync state relative to this script."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    full_path = os.path.join(base_dir, path)
    if os.path.exists(full_path):
        with open(full_path) as f:
            return json.load(f)
    return {}

def save_state(path: str, state: Dict[str, str]):
    """Save sync state relative to this script."""
    base_dir = os.path.dirname(os.path.abspath(__file__))
    full_path = os.path.join(base_dir, path)
    with open(full_path, "w") as f:
        json.dump(state, f, indent=2)

def get_existing_events(service, calendar_id: str, time_min: str, time_max: str) -> Dict[str, Dict[str, Any]]:
    """Retrieve existing events in the specified window."""
    existing_events: Dict[str, Dict[str, Any]] = {}
    page_token = None
    while True:
        res = gexec(
            service.events().list(
                calendarId=calendar_id,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                showDeleted=False,
                maxResults=2500,
                pageToken=page_token
            ),
            verb_hint="events.list"
        )
        for ev in res.get("items", []):
            existing_events[ev["id"]] = ev
        page_token = res.get("nextPageToken")
        if not page_token:
            break
    return existing_events

# ----------------------------------------------------------------------
# Main sync logic
# ----------------------------------------------------------------------
def safe_sync():
    cfg = load_config()
    calendar_id = cfg.get("google_calendar_id", "primary")
    tzname = cfg.get("timezone", "UTC")
    sync_days_past = cfg.get("sync_days_past", 90)
    sync_days_future = cfg.get("sync_days_future", 120)
    api_delay = cfg.get("api_delay_seconds", 0.05)

    tz = gettz(tzname)
    now = datetime.now(tz)
    start_window = now - timedelta(days=sync_days_past)
    end_window = now + timedelta(days=sync_days_future)

    # Prepare outbox
    base_dir = os.path.dirname(os.path.abspath(__file__))
    outbox = os.path.join(base_dir, "outbox")
    ics_files = [os.path.join(outbox, f) for f in os.listdir(outbox) if f.startswith("clean_") and f.endswith(".ics")]
    if not ics_files:
        log.error("No cleaned .ics files found. Run clean_ics_files.py first.")
        sys.exit(1)

    # Parse all .ics files
    all_instances: List[Dict[str, Any]] = []
    for path in ics_files:
        all_instances.extend(parse_ics(path, tzname, start_window, end_window))
    log.info(f"[INFO] Parsed {len(all_instances)} event instances from .ics files.")

    # Build new_events keyed by composite key
    new_events: Dict[str, Dict[str, Any]] = {}
    for inst in all_instances:
        key = f"{inst['uid']}|{inst['start'].isoformat()}"  # uid plus start ISO
        new_events[key] = inst

    # Load previous sync state
    state_file = "sync_state.json"
    state = load_state(state_file)
    service = get_service()

    # Fetch existing events within window
    time_min = start_window.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")
    time_max = end_window.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")
    existing_events = get_existing_events(service, calendar_id, time_min, time_max)

    created = updated = skipped = 0

    # Delete events no longer present
    for key in list(state.keys()):
        if key not in new_events:
            event_id = state[key]
            if event_id in existing_events:
                try:
                    gexec(service.events().delete(calendarId=calendar_id, eventId=event_id),
                          verb_hint="events.delete", uid_hint=key, delay=api_delay)
                    log.info(f"[DEL ] {key} -> event {event_id}")
                except HttpError as e:
                    if getattr(e.resp, "status", None) in (404, 410):
                        log.info(f"[DEL ] {key} already absent")
                    else:
                        log.error(f"[ERR ] Deleting {key}: {e}")
            state.pop(key, None)

    # Insert or update events
    for key, inst in new_events.items():
        event_id = state.get(key) or safe_event_id(key)
        body = {
            "summary": inst["summary"],
            "description": inst["description"],
            "location": inst["location"],
            "start": {"dateTime": inst["start"].isoformat()} if not inst["all_day"] else {"date": inst["start"].date().isoformat()},
            "end": {"dateTime": inst["end"].isoformat()} if not inst["all_day"] else {"date": inst["end"].date().isoformat()},
        }
        try:
            if event_id in existing_events:
                existing = existing_events[event_id]
                if (existing.get("summary") == body["summary"]
                        and existing.get("description") == body["description"]
                        and existing.get("location") == body["location"]
                        and existing.get("start") == body["start"]
                        and existing.get("end") == body["end"]):
                    skipped += 1
                    state[key] = event_id
                    continue
                gexec(service.events().patch(calendarId=calendar_id, eventId=event_id, body=body),
                      verb_hint="events.patch", uid_hint=key, delay=api_delay)
                updated += 1
            else:
                gexec(service.events().insert(calendarId=calendar_id, body={"id": event_id, **body}),
                      verb_hint="events.insert", uid_hint=key, delay=api_delay)
                created += 1
            state[key] = event_id
        except HttpError as e:
            if getattr(e.resp, "status", None) == 409:
                gexec(service.events().patch(calendarId=calendar_id, eventId=event_id, body=body),
                      verb_hint="events.patch", uid_hint=key, delay=api_delay)
                updated += 1
                state[key] = event_id
            else:
                raise

    save_state(state_file, state)
    log.info(f"[DONE] Created: {created}, Updated: {updated}, Skipped: {skipped}")

if __name__ == "__main__":
    safe_sync()
