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
- Backoff + jitter + steady throttle for Google API rate limits
"""
import argparse
import json
import os
import re
import sys
import time
import random
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

# How chatty we are with Google (base steady throttle + retry backoff on top)
STEADY_THROTTLE_SECONDS = 0.25  # ~4 calls/sec max steady state (safe for per-minute caps)
MAX_RETRIES = 8                 # retry budget per API call
INITIAL_BACKOFF = 1.0           # seconds
MAX_BACKOFF = 32.0              # cap a single sleep

BATCH_PROGRESS_STEP = 50

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
                print(f"[WARN] Skipping VCALENDAR chunk #{i}: {e}", file=sys.stderr)
        if cals:
            return cals
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

# --------------------
# Google API helpers
# --------------------
_last_call_ts = 0.0

def steady_throttle():
    """Simple steady throttle to keep under per-minute caps."""
    global _last_call_ts
    now = time.time()
    elapsed = now - _last_call_ts
    if elapsed < STEADY_THROTTLE_SECONDS:
        time.sleep(STEADY_THROTTLE_SECONDS - elapsed)
    _last_call_ts = time.time()

RETRYABLE_STATUS = {403, 429, 500, 502, 503, 504}

def _retry_after_seconds(e: HttpError) -> Optional[float]:
    try:
        retry_after = e.resp.get('retry-after') or e.resp.get('Retry-After')
        if retry_after:
            return float(retry_after)
    except Exception:
        pass
    return None

def g_execute(request_callable, *, op_desc: str = "", verbose: bool = True):
    """
    Execute a googleapiclient request with:
      - steady per-call throttle
      - exponential backoff with jitter on rate/server errors
      - honors Retry-After if present
    request_callable: a zero-arg lambda returning request.execute()
    """
    steady_throttle()
    delay = INITIAL_BACKOFF
    for attempt in range(1, MAX_RETRIES + 1):
        try:
            return request_callable()
        except HttpError as e:
            status = getattr(e.resp, "status", None)
            reason = getattr(e, "reason", str(e))
            # Non-retryable
            if status not in RETRYABLE_STATUS:
                if verbose:
                    print(f"[ERR ] {op_desc} failed (non-retryable {status}): {reason}", flush=True)
                raise
            # Retryable: compute sleep
            ra = _retry_after_seconds(e)
            sleep_for = ra if ra is not None else min(MAX_BACKOFF, delay * (1 + 0.25 * random.random()))
            if verbose:
                print(f"[rate] {op_desc} got {status}; sleeping {sleep_for:.1f}s (attempt {attempt}/{MAX_RETRIES})", flush=True)
            time.sleep(sleep_for)
            delay = min(MAX_BACKOFF, delay * 2)
        except Exception as ex:
            # Unknown error: one quick retry path then give up
            if attempt >= 2:
                if verbose:
                    print(f"[ERR ] {op_desc} failed: {ex}", flush=True)
                raise
            time.sleep(1.0)
    # If we fell out of loop, raise last HttpError again
    raise RuntimeError(f"{op_desc} exhausted retries")

def find_by_private_uid(service, calendar_id: str, uid: str, verbose: bool = True) -> Optional[Dict[str, Any]]:
    """Search Google by our private extended property icalUID=<uid>."""
    try:
        res = g_execute(
            lambda: service.events().list(
                calendarId=calendar_id,
                privateExtendedProperty=f"icalUID={uid}",
                maxResults=5,
                singleEvents=False
            ).execute(),
            op_desc=f"events.list privateExtendedProperty for uid={uid}",
            verbose=verbose
        )
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
            updated = g_execute(
                lambda: service.events().patch(calendarId=calendar_id, eventId=mapped_id, body=body).execute(),
                op_desc=f"events.patch uid={uid}",
                verbose=verbose
            )
            return ("update", updated["id"])
        except HttpError as e:
            if e.resp.status not in (404, 410):
                raise
            # fall through if not found

    found = find_by_private_uid(service, calendar_id, uid, verbose=verbose)
    if found:
        try:
            updated = g_execute(
                lambda: service.events().patch(calendarId=calendar_id, eventId=found["id"], body=body).execute(),
                op_desc=f"events.patch(found) uid={uid}",
                verbose=verbose
            )
            state[uid] = updated["id"]
            return ("update", updated["id"])
        except HttpError as e:
            if e.resp.status not in (404, 410):
                raise

    created = g_execute(
        lambda: service.events().insert(calendarId=calendar_id, body=body).execute(),
        op_desc=f"events.insert uid={uid}",
        verbose=verbose
    )
    state[uid] = created["id"]
    return ("create", created["id"])

def current_index(service, calendar_id: str, time_min: Optional[str], time_max: Optional[str], verbose: bool) -> Dict[str, Dict[str, Any]]:
    """
    Build a lightweight index of current Google events keyed by our private icalUID.
    We fetch events within the same window and then pick those that actually
    have extendedProperties.private.icalUID.
    """
    idx: Dict[str, Dict[str, Any]] = {}
    page_token = None
    page = 0
    while True:
        kwargs = dict(
            calendarId=calendar_id,
            maxResults=500,         # keep page small -> fewer per-call payloads, easier on quota
            pageToken=page_token,
            singleEvents=False,
            showDeleted=False,
        )
        if time_min:
            kwargs["timeMin"] = time_min
        if time_max:
            kwargs["timeMax"] = time_max

        res = g_execute(
            lambda: service.events().list(**kwargs).execute(),
            op_desc=f"events.list page={page}",
            verbose=verbose
        )
        for e in res.get("items", []):
            priv = (e.get("extendedProperties") or {}).get("private") or {}
            uid = priv.get("icalUID")
            if uid:
                idx[uid] = e
        page_token = res.get("nextPageToken")
        page += 1
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

    # Parse ICS for desired instances
    instances = parse_ics_instances(ICS_PATH, tzname, past_days, future_days, verbose)
    log(f"Parsed {len(instances)} event instances from ICS window.", verbose=verbose)

    # Compute the same window in RFC3339 UTC for Google list calls
    tz = pytz.timezone(tzname)
    g_time_min = (datetime.now(tz) - timedelta(days=past_days)).astimezone(timezone.utc).isoformat()
    g_time_max = (datetime.now(tz) + timedelta(days=future_days)).astimezone(timezone.utc).isoformat()

    # Create/update pass
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
            # Already retried; log a concise line and keep going
            reason = getattr(e, "reason", str(e))
            print(f"[ERR ] Upsert failed for '{inst['summary']}' uid={inst['uid']}: {reason}", flush=True)
        if i % BATCH_PROGRESS_STEP == 0:
            log(f"[progress] processed {i}/{len(instances)}", verbose=verbose)
            # light breather every chunk (helps minute-based caps)
            time.sleep(1.0)
    save_state(state)

    # Delete pass: remove Google events (with our icalUID) that are no longer in the ICS window
    current = current_index(service, calendar_id, g_time_min, g_time_max, verbose)
    deletes = 0
    desired_uids = {i["uid"] for i in instances}
    for j, (uid, ge) in enumerate(current.items(), start=1):
        if uid not in desired_uids:
            try:
                g_execute(
                    lambda: service.events().delete(calendarId=calendar_id, eventId=ge["id"]).execute(),
                    op_desc=f"events.delete uid={uid}",
                    verbose=verbose
                )
                deletes += 1
                state.pop(uid, None)
            except HttpError as e:
                if e.resp.status not in (404, 410):
                    print(f"[WARN] Delete failed for uid={uid}: {e}", flush=True)
        if j % 100 == 0:
            time.sleep(0.5)
    save_state(state)

    dt = time.time() - t0
    log(f"Sync done: created={creates}, updated={updates}, deleted={deletes} in {dt:.1f}s", verbose=verbose)

if __name__ == "__main__":
    main()
