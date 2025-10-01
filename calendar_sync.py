#!/usr/bin/env python3
import os, sys, json, hashlib, random, time, re, logging
from datetime import datetime, timedelta
from typing import Dict, Any, List
from dateutil.tz import gettz
import recurring_ical_events as rie
from icalendar import Calendar
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# =============================================================================
# Logging
# =============================================================================
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    handlers=[logging.StreamHandler(sys.stdout)],
)
log = logging.getLogger("calendar_sync")

# =============================================================================
# Utils
# =============================================================================
def safe_event_id(uid: str) -> str:
    """Generate a Google-compliant deterministic event ID from UID."""
    digest = hashlib.sha1(uid.encode("utf-8")).hexdigest()[:32]
    event_id = f"e{digest}"
    return event_id

def _as_tz(dt: datetime, tzname: str) -> datetime:
    tz = gettz(tzname)
    return dt.astimezone(tz)

def _steady_throttle():
    time.sleep(0.05)

def _is_rate_limited_error(e: Exception) -> bool:
    try:
        status = getattr(e, "resp", None).status if hasattr(e, "resp") else None
        if status in (403, 429):
            return True
        content = getattr(e, "content", b"") or b""
        if status == 400 and (b"rateLimitExceeded" in content or b"quota" in content):
            return True
        return False
    except Exception:
        return False

def gexec(request_callable, verb_hint: str = "", uid_hint: str = ""):
    attempt = 0
    while True:
        _steady_throttle()
        try:
            return request_callable.execute()
        except HttpError as e:
            attempt += 1
            if _is_rate_limited_error(e) and attempt <= 5:
                sleep_s = min(2.0 * attempt, 6.0) + random.uniform(0.1, 0.5)
                if verb_hint and uid_hint:
                    log.info(f"[rate] {verb_hint} uid={uid_hint} got {e.resp.status}; sleeping {sleep_s:.1f}s (attempt {attempt}/5)")
                else:
                    log.info(f"[rate] got {e.resp.status}; sleeping {sleep_s:.1f}s (attempt {attempt}/5)")
                time.sleep(sleep_s)
                continue
            raise

# =============================================================================
# Config / Auth
# =============================================================================
def load_config() -> Dict[str, Any]:
    with open("calendar_config.json") as f:
        return json.load(f)

def get_service():
    creds = Credentials.from_authorized_user_file("token.json", ["https://www.googleapis.com/auth/calendar"])
    return build("calendar", "v3", credentials=creds, cache_discovery=False)

# =============================================================================
# Main Sync
# =============================================================================
def parse_ics(path: str, tzname: str, start_window: datetime, end_window: datetime) -> List[Dict[str, Any]]:
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
        if comp.name not in ("VEVENT",):
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
            dtstart = _as_tz(dtstart, tzname)
        else:
            all_day = True
            dtstart = tz.localize(datetime(dtstart.year, dtstart.month, dtstart.day))
        if isinstance(dtend, datetime):
            dtend = _as_tz(dtend, tzname)
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

def sync_events(service, calendar_id: str, instances: List[Dict[str, Any]]):
    kwargs = {
        "calendarId": calendar_id,
        "timeMin": (datetime.utcnow() - timedelta(days=30)).isoformat() + "Z",
        "timeMax": (datetime.utcnow() + timedelta(days=365)).isoformat() + "Z",
        "singleEvents": True,
        "showDeleted": False,
    }
    existing = gexec(service.events().list(**kwargs), verb_hint="events.list")
    existing_items = existing.get("items", [])
    existing_index = {e.get("id"): e for e in existing_items}
    log.info(f"[INFO] Indexed {len(existing_index)} Google events in window")

    created = 0
    updated = 0
    skipped = 0

    for inst in instances:
        event_id = safe_event_id(inst["uid"])
        body = {
            "summary": inst["summary"],
            "description": inst["description"],
            "location": inst["location"],
            "start": {"dateTime": inst["start"].isoformat()} if not inst["all_day"] else {"date": inst["start"].date().isoformat()},
            "end": {"dateTime": inst["end"].isoformat()} if not inst["all_day"] else {"date": inst["end"].date().isoformat()},
        }
        
        try:
            if event_id in existing_index:
                existing = existing_index[event_id]
                # Compare - only patch if changed
                if (existing.get("summary") == body["summary"] and
                    existing.get("description") == body["description"] and
                    existing.get("location") == body["location"] and
                    existing.get("start") == body["start"] and
                    existing.get("end") == body["end"]):
                    skipped += 1
                    continue
                
                gexec(
                    service.events().patch(
                        calendarId=calendar_id,
                        eventId=event_id,
                        body=body,
                    ),
                    verb_hint="events.patch",
                    uid_hint=inst["uid"],
                )
                updated += 1
                continue
            
            gexec(
                service.events().insert(
                    calendarId=calendar_id,
                    body={"id": event_id, **body},
                ),
                verb_hint="events.insert",
                uid_hint=inst["uid"],
            )
            created += 1
        except HttpError as e:
            if getattr(e, "resp", None) and e.resp.status == 404:
                gexec(
                    service.events().insert(
                        calendarId=calendar_id,
                        body={"id": event_id, **body},
                    ),
                    verb_hint="events.insert",
                    uid_hint=inst["uid"],
                )
                created += 1
            elif getattr(e, "resp", None) and e.resp.status == 409:
                gexec(
                    service.events().patch(
                        calendarId=calendar_id,
                        eventId=event_id,
                        body=body,
                    ),
                    verb_hint="events.patch",
                    uid_hint=inst["uid"],
                )
                updated += 1
            else:
                raise
    
    log.info(f"[DONE] Created: {created}, Updated: {updated}, Skipped: {skipped}")

# =============================================================================
# Entrypoint
# =============================================================================
def main():
    cfg = load_config()
    calendar_id = cfg.get("google_calendar_id", "primary")
    tzname = cfg.get("timezone", "America/New_York")
    start_window = datetime.now() - timedelta(days=30)
    end_window = datetime.now() + timedelta(days=365)

    # Check if source changed since last sync
    source_ics = "outbox/outlook_full_export.ics"
    state_file = "last_sync_time.txt"
    
    if os.path.exists(state_file) and os.path.exists(source_ics):
        with open(state_file) as f:
            last_sync = float(f.read().strip())
        
        if os.path.getmtime(source_ics) < last_sync:
            log.info("[SKIP] No changes to Outlook export since last sync")
            sys.exit(0)

    outbox = os.path.join("outbox")
    ics_files = [os.path.join(outbox, f) for f in os.listdir(outbox) if f.startswith("clean_") and f.endswith(".ics")]
    if not ics_files:
        log.error("No cleaned ICS files found. Did you run clean_ics_files.py?")
        sys.exit(1)

    all_instances = []
    for path in ics_files:
        insts = parse_ics(path, tzname, start_window, end_window)
        all_instances.extend(insts)
    log.info(f"Parsed {len(all_instances)} event instances from ICS files.")

    service = get_service()
    sync_events(service, calendar_id, all_instances)
    
    # Mark successful sync
    with open(state_file, 'w') as f:
        f.write(str(time.time()))

if __name__ == "__main__":
    main()
