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

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("safe_sync")

def safe_event_id(key: str) -> str:
    """Generate a deterministic Google event ID from a composite key."""
    digest = hashlib.sha1(key.encode("utf-8")).hexdigest()[:32]
    return f"e{digest}"

def gexec(request_callable, verb_hint: str = "", uid_hint: str = "", delay: float = 0.05):
    """Execute a Google API request with retry on rate limit errors."""
    attempt = 0
    while True:
        time.sleep(delay)
        try:
            return request_callable.execute()
        except HttpError as e:
            attempt += 1
            if e.resp.status in (403, 429) and attempt <= 5:
                sleep_s = min(2.0 * attempt, 6.0) + random.uniform(0.1, 0.5)
                log.info(f"[rate] {verb_hint} uid={uid_hint} got {e.resp.status}; sleeping {sleep_s:.1f}s (attempt {attempt}/5)")
                time.sleep(sleep_s)
                continue
            raise

def parse_ics(path: str, tzname: str, start_window: datetime, end_window: datetime) -> List[Dict[str, Any]]:
    """Parse a cleaned .ics file into a list of event dicts within the window."""
    tz = gettz(tzname)
    instances = []
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
    with open("calendar_config.json") as f:
        return json.load(f)

def get_service():
    SCOPES = ["https://www.googleapis.com/auth/calendar"]
    creds = None
    if os.path.exists("token.json"):
        creds = Credentials.from_authorized_user_file("token.json", SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception:
                pass
        if not creds or not creds.valid:
            flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
            creds = flow.run_local_server(port=0)
        with open("token.json", "w") as token:
            token.write(creds.to_json())
    return build("calendar", "v3", credentials=creds, cache_discovery=False)

def load_state(path: str) -> Dict[str, str]:
    if os.path.exists(path):
        with open(path) as f:
            return json.load(f)
    return {}

def save_state(path: str, state: Dict[str, str]):
    with open(path, "w") as f:
        json.dump(state, f, indent=2)

def get_existing_events(service, calendar_id: str, time_min: str, time_max: str) -> Dict[str, Dict[str, Any]]:
    """Retrieve all Google Calendar events in the given time range and return a dict keyed by event ID."""
    existing_events = {}
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
            verb_hint="events.list",
        )
        for ev in res.get("items", []):
            existing_events[ev["id"]] = ev
        page_token = res.get("nextPageToken")
        if not page_token:
            break
    return existing_events

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

    # Gather instances
    outbox = "outbox"
    ics_files = [os.path.join(outbox, f) for f in os.listdir(outbox) if f.startswith("clean_") and f.endswith(".ics")]
    if not ics_files:
        log.error("No cleaned .ics files found. Run clean_ics_files.py first.")
        sys.exit(1)

    all_instances: List[Dict[str, Any]] = []
    for path in ics_files:
        all_instances.extend(parse_ics(path, tzname, start_window, end_window))
    log.info(f"[INFO] Parsed {len(all_instances)} event instances from .ics files.")

    # Build a dict keyed by composite key (UID + start) so each instance is unique
    new_events: Dict[str, Dict[str, Any]] = {}
    for inst in all_instances:
        key = f"{inst['uid']}|{inst['start'].isoformat()}"
        new_events[key] = inst

    # Load state mapping from composite key -> Google event ID
    state_file = "sync_state.json"
    state = load_state(state_file)

    service = get_service()

    # Prepare timeMin/timeMax in UTC
    time_min = start_window.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")
    time_max = end_window.astimezone(timezone.utc).isoformat().replace("+00:00", "Z")
    existing_events = get_existing_events(service, calendar_id, time_min, time_max)

    created = 0
    updated = 0
    skipped = 0

    # Deletions: remove any events that were previously synced but are no longer present
    for key in list(state.keys()):
        if key not in new_events:
            event_id = state[key]
            if event_id in existing_events:
                try:
                    gexec(
                        service.events().delete(calendarId=calendar_id, eventId=event_id),
                        verb_hint="events.delete",
                        uid_hint=key,
                        delay=api_delay,
                    )
                    log.info(f"[DEL ] {key} -> event {event_id}")
                except HttpError as e:
                    if getattr(e, "resp", None) and e.resp.status in (404, 410):
                        log.info(f"[DEL ] {key} already absent")
                    else:
                        log.error(f"[ERR ] Deleting {key}: {e}")
            state.pop(key, None)

    # Insert/update each event instance
    for key, inst in new_events.items():
        # Generate deterministic event ID from composite key
        event_id = state.get(key) or safe_event_id(key)
        body = {
            "summary": inst["summary"],
            "description": inst["description"],
            "location": inst["location"],
            "start": (
                {"dateTime": inst["start"].isoformat()}
                if not inst["all_day"]
                else {"date": inst["start"].date().isoformat()}
            ),
            "end": (
                {"dateTime": inst["end"].isoformat()}
                if not inst["all_day"]
                else {"date": inst["end"].date().isoformat()}
            ),
        }
        try:
            if event_id in existing_events:
                existing = existing_events[event_id]
                if (
                    existing.get("summary") == body["summary"]
                    and existing.get("description") == body["description"]
                    and existing.get("location") == body["location"]
                    and existing.get("start") == body["start"]
                    and existing.get("end") == body["end"]
                ):
                    skipped += 1
                    state[key] = event_id
                    continue
                # Update
                gexec(
                    service.events().patch(
                        calendarId=calendar_id, eventId=event_id, body=body
                    ),
                    verb_hint="events.patch",
                    uid_hint=key,
                    delay=api_delay,
                )
                updated += 1
            else:
                # Create
                gexec(
                    service.events().insert(
                        calendarId=calendar_id, body={"id": event_id, **body}
                    ),
                    verb_hint="events.insert",
                    uid_hint=key,
                    delay=api_delay,
                )
                created += 1
            state[key] = event_id
        except HttpError as e:
            if getattr(e, "resp", None) and e.resp.status == 409:
                gexec(
                    service.events().patch(
                        calendarId=calendar_id, eventId=event_id, body=body
                    ),
                    verb_hint="events.patch",
                    uid_hint=key,
                    delay=api_delay,
                )
                updated += 1
                state[key] = event_id
            else:
                raise

    save_state(state_file, state)
    log.info(f"[DONE] Created: {created}, Updated: {updated}, Skipped: {skipped}")

if __name__ == "__main__":
    safe_sync()
