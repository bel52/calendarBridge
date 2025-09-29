#!/usr/bin/env python3
"""
CalendarBridge: Outlook (ICS) -> Google Calendar (one-way)

Key behaviors:
- Parse ~/calendarBridge/outbox/outlook_full_export.ics
- Supports files that contain multiple VCALENDAR sections (concatenated exports)
- Expand recurring VEVENTs into instances (limited window)
- Dedupe by ICS UID using extendedProperties.private.icalUID
- Never set 'id' on insert (avoid 409 duplicate/conflict)
- Maintain sync_state.json mapping { icalUID: google_event_id }
- If mapping is missing, search by privateExtendedProperty
- Respect date window (default: past 30d -> next 365d; configurable)
- Verbose logging with --verbose
"""
import argparse
import json
import os
import re
import sys
import time
from datetime import datetime, timedelta, timezone
from typing import Dict, Any, List, Tuple, Optional

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request

import pytz
from icalendar import Calendar
import recurring_ical_events

ROOT = os.path.abspath(os.path.dirname(__file__))
OUTBOX = os.path.join(ROOT, "outbox")
ICS_PATH = os.path.join(OUTBOX, "outlook_full_export.ics")
STATE_PATH = os.path.join(ROOT, "sync_state.json")
CONFIG_PATH = os.path.join(ROOT, "calendar_config.json")

SCOPES = ["https://www.googleapis.com/auth/calendar"]

DEFAULT_PAST_DAYS = 30
DEFAULT_FUTURE_DAYS = 365
BATCH_SIZE = 50

VCAL_RE = re.compile(br"BEGIN:VCALENDAR\r?\n.*?END:VCALENDAR\r?\n?", re.DOTALL)

def log(msg: str, *, verbose: bool = True):
    if verbose:
        print(msg, flush=True)

def load_config() -> Dict[str, Any]:
    cfg = {
        "target_calendar_id": "primary",
        "tz": "America/New_York",
        "past_days": DEFAULT_PAST_DAYS,
        "future_days": DEFAULT_FUTURE_DAYS,
    }
    try:
        with open(CONFIG_PATH, "r") as f:
            user = json.load(f)
        cfg.update({k: v for k, v in user.items() if v is not None})
    except FileNotFoundError:
        pass
    return cfg

def get_service() -> Any:
    cred_path = os.path.join(ROOT, "credentials.json")
    token_path = os.path.join(ROOT, "token.json")
    if not os.path.exists(cred_path):
        raise FileNotFoundError(f"credentials.json not found at {cred_path}")

    creds = None
    if os.path.exists(token_path):
        creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise RuntimeError("Missing/invalid token.json. Re-run to authorize.")
    return build("calendar", "v3", credentials=creds, cache_discovery=False)

def load_state() -> Dict[str, str]:
    try:
        with open(STATE_PATH, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return {}
    except json.JSONDecodeError:
        return {}

def save_state(state: Dict[str, str]):
    tmp = STATE_PATH + ".tmp"
    with open(tmp, "w") as f:
        json.dump(state, f, indent=2, sort_keys=True)
    os.replace(tmp, STATE_PATH)

def read_calendars_from_bytes(data: bytes) -> List[Calendar]:
    """
    Return a list of icalendar.Calendar objects from raw bytes.
    Handles concatenated VCALENDAR sections.
    """
    chunks = VCAL_RE.findall(data)
    if chunks:
        cals = []
        for i, chunk in enumerate(chunks, 1):
            try:
                cals.append(Calendar.from_ical(chunk))
            except Exception as e:
                # Skip bad chunks but report
                print(f"[WARN] Skipping VCALENDAR chunk #{i}: {e}", file=sys.stderr)
        if cals:
            return cals
        # Fall through to single parse if none worked
    # Single calendar case
    return [Calendar.from_ical(data)]

def parse_ics_instances(ics_path: str, tzname: str, past_days: int, future_days: int, verbose: bool) -> List[Dict[str, Any]]:
    if not os.path.exists(ics_path):
        raise FileNotFoundError(f"ICS export file not found at {ics_path}")

    with open(ics_path, "rb") as f:
        raw = f.read()

    calendars = read_calendars_from_bytes(raw)

    tz = pytz.timezone(tzname)
    start = datetime.now(tz) - timedelta(days=past_days)
    end = datetime.now(tz) + timedelta(days=future_days)

    instances: List[Dict[str, Any]] = []
    for cal in calendars:
        # Expand recurrences within the window
        events = recurring_ical_events.of(cal).between(start, end)
        for e in events:
            uid = str(e.get("UID"))
            if not uid or uid.strip() == "":
                continue

            summary = str(e.get("SUMMARY") or "").strip()
            description = str(e.get("DESCRIPTION") or "").strip()
            location = str(e.get("LOCATION") or "").strip()

            dtstart = e.get("DTSTART").dt
            dtend = e.get("DTEND").dt if e.get("DTEND") else None

            # Normalize to Google format
            if hasattr(dtstart, "hour"):
                if dtstart.tzinfo is None:
                    dtstart = tz.localize(dtstart)
                start_body = {"dateTime": dtstart.astimezone(timezone.utc).isoformat()}
            else:
                start_body = {"date": dtstart.isoformat()}

            if dtend is None:
                if "dateTime" in start_body:
                    end_dt = (datetime.fromisoformat(start_body["dateTime"].replace("Z", "+00:00")) + timedelta(hours=1)).isoformat()
                    end_body = {"dateTime": end_dt}
                else:
                    end_body = {"date": start_body["date"]}
            else:
                if hasattr(dtend, "hour"):
                    if dtend.tzinfo is None:
                        dtend = tz.localize(dtend)
                    end_body = {"dateTime": dtend.astimezone(timezone.utc).isoformat()}
                else:
                    end_body = {"date": dtend.isoformat()}

            instances.append({
                "uid": uid,
                "summary": summary or "(no title)",
                "description": description,
                "location": location,
                "start": start_body,
                "end": end_body,
            })

    return instances

def find_by_private_uid(service, calendar_id: str, uid: str) -> Optional[Dict[str, Any]]:
    try:
        res = service.events().list(
            calendarId=calendar_id,
            privateExtendedProperty=f"icalUID={uid}",
            maxResults=5,
            singleEvents=False
        ).execute()
        items = res.get("items", [])
        return items[0] if items else None
    except HttpError:
        return None

def upsert_event(service, calendar_id: str, state: Dict[str, str], inst: Dict[str, Any], verbose: bool) -> Tuple[str, str]:
    uid = inst["uid"]
    body = {
        "summary": inst["summary"],
        "location": inst["location"] or None,
        "description": inst["description"] or None,
        "start": inst["start"],
        "end": inst["end"],
        "extendedProperties": {
            "private": {
                "icalUID": uid
            }
        }
    }

    mapped_id = state.get(uid)
    if mapped_id:
        try:
            updated = service.events().patch(calendarId=calendar_id, eventId=mapped_id, body=body).execute()
            return ("update", updated["id"])
        except HttpError as e:
            if e.resp.status not in (404, 410):
                raise

    found = find_by_private_uid(service, calendar_id, uid)
    if found:
        try:
            updated = service.events().patch(calendarId=calendar_id, eventId=found["id"], body=body).execute()
            state[uid] = updated["id"]
            return ("update", updated["id"])
        except HttpError as e:
            if e.resp.status not in (404, 410):
                raise

    created = service.events().insert(calendarId=calendar_id, body=body).execute()
    state[uid] = created["id"]
    return ("create", created["id"])

def current_index(service, calendar_id: str, verbose: bool) -> Dict[str, Dict[str, Any]]:
    idx: Dict[str, Dict[str, Any]] = {}
    page_token = None
    while True:
        res = service.events().list(
            calendarId=calendar_id,
            maxResults=2500,
            pageToken=page_token,
            singleEvents=False,
            showDeleted=False,
            privateExtendedProperty="icalUID"
        ).execute()
        for e in res.get("items", []):
            priv = (e.get("extendedProperties") or {}).get("private") or {}
            uid = priv.get("icalUID")
            if uid:
                idx[uid] = e
        page_token = res.get("nextPageToken")
        if not page_token:
            break
    return idx

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--verbose", action="store_true", help="verbose logs")
    ap.add_argument("--past-days", type=int, default=None)
    ap.add_argument("--future-days", type=int, default=None)
    args = ap.parse_args()
    verbose = args.verbose

    cfg = load_config()
    calendar_id = cfg.get("target_calendar_id", "primary")
    tzname = cfg.get("tz", "America/New_York")
    past_days = args.past_days if args.past_days is not None else int(cfg.get("past_days", DEFAULT_PAST_DAYS))
    future_days = args.future_days if args.future_days is not None else int(cfg.get("future_days", DEFAULT_FUTURE_DAYS))

    log("--- Starting CalendarBridge Sync Cycle ---", verbose=verbose)
    log(f"Config: calendar_id={calendar_id} tz={tzname} window=[-{past_days}d,+{future_days}d]", verbose=verbose)

    if not os.path.exists(ICS_PATH):
        print(f"[ERR ] Export file not found: {ICS_PATH}", file=sys.stderr)
        sys.exit(2)

    service = get_service()
    state = load_state()

    instances = parse_ics_instances(ICS_PATH, tzname, past_days, future_days, verbose)
    log(f"Parsed {len(instances)} event instances from ICS window.", verbose=verbose)

    creates = updates = 0
    t0 = time.time()
    for i, inst in enumerate(instances, start=1):
        try:
            action, eid = upsert_event(service, calendar_id, state, inst, verbose)
            if action == "create":
                creates += 1
            else:
                updates += 1
        except HttpError as e:
            reason = getattr(e, "reason", str(e))
            print(f"[ERR ] Upsert failed for '{inst['summary']}' uid={inst['uid']}: {reason}", flush=True)
        if i % BATCH_SIZE == 0:
            log(f"[progress] processed {i}/{len(instances)}", verbose=verbose)
    save_state(state)

    current = current_index(service, calendar_id, verbose)
    deletes = 0
    desired_uids = {i["uid"] for i in instances}
    for uid, ge in current.items():
        if uid not in desired_uids:
            try:
                service.events().delete(calendarId=calendar_id, eventId=ge["id"]).execute()
                deletes += 1
                state.pop(uid, None)
            except HttpError as e:
                if e.resp.status not in (404, 410):
                    print(f"[WARN] Delete failed for uid={uid}: {e}", flush=True)
    save_state(state)

    dt = time.time() - t0
    log(f"Sync done: created={creates}, updated={updates}, deleted={deletes} in {dt:.1f}s", verbose=verbose)

if __name__ == "__main__":
    main()
