#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, sys, json, time, random, glob, math
from typing import Dict, Any, Optional, Tuple, List

# ---------------- Config & helpers ----------------

def _cfg_load(path: str) -> Dict[str, Any]:
    with open(path, "r") as f:
        return json.load(f)

def _sleep_with_jitter(base_seconds: float) -> None:
    if base_seconds <= 0:
        return
    jitter = base_seconds * random.uniform(-0.20, 0.20)
    time.sleep(max(0.0, base_seconds + jitter))

def _get_retry_after(ex: Exception) -> Optional[float]:
    try:
        headers = getattr(ex, "resp", None).headers or {}
    except Exception:
        headers = {}
    ra = headers.get("Retry-After")
    if not ra:
        return None
    try:
        return float(ra)
    except Exception:
        return None

ROOT = os.path.expanduser("~/calendarBridge")
CONFIG_PATH = os.path.join(ROOT, "calendar_config.json")
CONF = _cfg_load(CONFIG_PATH) if os.path.exists(CONFIG_PATH) else {}

API_DELAY = float(CONF.get("api_delay_seconds", 1.05))
TIMEZONE  = CONF.get("timezone", "America/New_York")
CAL_ID    = CONF.get("google_calendar_id", "primary")
DAYS_PAST = int(CONF.get("sync_days_past", 60))
DAYS_FUT  = int(CONF.get("sync_days_future", 90))
DELETE_ORPHANS = True  # per user instruction
QUOTA_USER = os.environ.get("CALBRIDGE_QUOTA_USER")
SLOW_START_MS = int(os.environ.get("CALBRIDGE_SLOW_START_MS", "0"))

if SLOW_START_MS > 0:
    time.sleep(SLOW_START_MS / 1000.0)

# Late imports so CONFIG is ready
from datetime import datetime, timedelta, timezone
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from icalendar import Calendar

# OAuth
creds_path = os.path.join(ROOT, "token.json")
creds = Credentials.from_authorized_user_file(creds_path)
service = build("calendar", "v3", credentials=creds, cache_discovery=False)

def request_with_backoff(req_builder, op_desc: str, max_attempts: int = 6) -> Any:
    attempt = 0
    backoff = 1.0
    while True:
        attempt += 1
        try:
            _sleep_with_jitter(API_DELAY)
            req = req_builder()
            if QUOTA_USER:
                req.uri += ("&" if "?" in req.uri else "?") + f"quotaUser={QUOTA_USER}"
            return req.execute()
        except HttpError as e:
            msg = getattr(e, "content", b"")
            msg_str = msg.decode("utf-8", errors="ignore") if isinstance(msg, (bytes, bytearray)) else str(msg)
            is_rate = (e.resp.status == 403 and ("rateLimitExceeded" in msg_str or "userRateLimitExceeded" in msg_str))
            ra = _get_retry_after(e)
            if is_rate and attempt < max_attempts:
                wait = ra if ra is not None else backoff + random.uniform(0.0, 0.5)
                print('Encountered 403 Forbidden with reason "rateLimitExceeded"')
                print(f"[rate] {op_desc} got 403; sleeping {round(wait,2)}s (attempt {attempt}/{max_attempts})")
                time.sleep(wait)
                backoff = min(backoff * 2.0, 16.0)
                continue
            raise
        except Exception:
            if attempt < max_attempts:
                wait = backoff + random.uniform(0.0, 0.5)
                print(f"[warn] {op_desc} transient error; sleeping {round(wait,2)}s (attempt {attempt}/{max_attempts})")
                time.sleep(wait)
                backoff = min(backoff * 2.0, 16.0)
                continue
            raise

# ---------------- ICS ingest ----------------

def _iso(dt) -> str:
    # Force ISO 8601 with timezone when possible
    if hasattr(dt, "tzinfo") and dt.tzinfo is not None:
        return dt.isoformat()
    # Assume local TZ if naive
    return dt.replace(tzinfo=timezone.utc).isoformat()

def load_local(outbox_dir: str) -> Dict[str, Dict[str, Any]]:
    """Return dict keyed by (uid|startIso) for instances."""
    events: Dict[str, Dict[str, Any]] = {}
    paths = glob.glob(os.path.join(outbox_dir, "*.ics"))
    if not paths:
        # fallback to the combined export filename
        p = os.path.join(outbox_dir, "outlook_full_export.ics")
        if os.path.exists(p):
            paths = [p]

    for path in paths:
        try:
            with open(path, "rb") as f:
                cal = Calendar.from_ical(f.read())
            for comp in cal.walk("VEVENT"):
                uid = str(comp.get("UID"))
                dtstart = comp.get("DTSTART").dt
                dtend   = comp.get("DTEND").dt
                all_day = not hasattr(dtstart, "hour")
                key = f"{uid}|{_iso(dtstart)}"

                events[key] = {
                    "uid": uid,
                    "summary": (comp.get("SUMMARY") or "").strip(),
                    "location": (comp.get("LOCATION") or "").strip(),
                    "description": (comp.get("DESCRIPTION") or "").strip(),
                    "start": dtstart,
                    "end": dtend,
                    "allDay": all_day
                }
        except Exception as e:
            print(f"[warn] Failed to parse {path}: {e}")
    return events

def to_body(ev: Dict[str, Any], tz: str) -> Dict[str, Any]:
    if ev["allDay"]:
        return {
            "summary": ev["summary"] or None,
            "location": ev["location"] or None,
            "description": ev["description"] or None,
            "start": {"date": ev["start"].date().isoformat(), "timeZone": tz},
            "end":   {"date": ev["end"].date().isoformat(),   "timeZone": tz},
            # keep a private echo of the UID as a second anchor
            "extendedProperties": {"private": {"icalUID": ev["uid"]}}
        }
    else:
        return {
            "summary": ev["summary"] or None,
            "location": ev["location"] or None,
            "description": ev["description"] or None,
            "start": {"dateTime": _iso(ev["start"]), "timeZone": tz},
            "end":   {"dateTime": _iso(ev["end"]),   "timeZone": tz},
            "extendedProperties": {"private": {"icalUID": ev["uid"]}}
        }

def same_enough(desired: Dict[str, Any], current: Dict[str, Any]) -> bool:
    for k in ("summary", "description", "location"):
        if (desired.get(k) or "") != (current.get(k) or ""):
            return False
    def norm(t):
        if not t: return ""
        return t.get("dateTime") or (t.get("date") + "T00:00:00")
    return norm(desired.get("start")) == norm(current.get("start")) and \
           norm(desired.get("end"))   == norm(current.get("end"))

# ---------------- Google side ----------------

def time_window() -> Tuple[str, str]:
    now = datetime.now(timezone.utc)
    tmin = now - timedelta(days=DAYS_PAST)
    tmax = now + timedelta(days=DAYS_FUT)
    return tmin.isoformat(), tmax.isoformat()

def list_google_events(calendar_id: str, tmin: str, tmax: str) -> List[Dict[str, Any]]:
    items: List[Dict[str, Any]] = []
    page_token = None
    while True:
        def _builder():
            return service.events().list(
                calendarId=calendar_id,
                timeMin=tmin, timeMax=tmax,
                singleEvents=True, showDeleted=False,
                maxResults=2500, pageToken=page_token, orderBy="startTime"
            )
        resp = request_with_backoff(_builder, "events.list window")
        items.extend(resp.get("items", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    return items

def key_for_google_item(g: Dict[str, Any]) -> Optional[str]:
    # Prefer Google's native iCalUID; fall back to our private property.
    ical = g.get("iCalUID")
    if not ical:
        ical = (g.get("extendedProperties", {}).get("private", {}) or {}).get("icalUID")
    if not ical:
        return None
    st = g.get("start", {})
    start_iso = st.get("dateTime") or (st.get("date") + "T00:00:00")
    return f"{ical}|{start_iso}"

def find_by_icaluid(calendar_id: str, ical_uid: str) -> Optional[Dict[str, Any]]:
    try:
        resp = request_with_backoff(
            lambda: service.events().list(calendarId=calendar_id, iCalUID=ical_uid, singleEvents=True, maxResults=5),
            f"events.list iCalUID={ical_uid}"
        )
        items = resp.get("items", [])
        return items[0] if items else None
    except Exception as e:
        print(f"[warn] list by iCalUID failed for {ical_uid}: {e}")
        return None

# ---------------- Main sync ----------------

def main():
    outbox = os.path.join(ROOT, "outbox")
    local = load_local(outbox)
    print(f"[INFO] Parsed {len(local)} event instances from .ics files.")

    tmin, tmax = time_window()
    g_items = list_google_events(CAL_ID, tmin, tmax)

    # Build google map
    g_map: Dict[str, Dict[str, Any]] = {}
    for g in g_items:
        k = key_for_google_item(g)
        if k:
            g_map[k] = g

    created = updated = skipped = deleted = failed = 0

    # 1) INSERT or UPDATE
    for k, ev in local.items():
        uid = ev["uid"]
        body = to_body(ev, TIMEZONE)

        g = g_map.get(k)
        if g is None:
            # brand-new -> IMPORT so we can set iCalUID
            def _builder():
                b = dict(body)
                b["iCalUID"] = uid  # events.import honors this
                return service.events().import_(calendarId=CAL_ID, body=b)
            try:
                resp = request_with_backoff(_builder, f"events.import uid={uid}")
                if resp and resp.get("id"):
                    created += 1
                else:
                    failed += 1
            except Exception as e:
                print(f"[warn] import failed uid={uid}: {e}")
                failed += 1
            continue

        # Exists -> PATCH if changed
        current_reduced = {
            "summary": g.get("summary"),
            "description": g.get("description"),
            "location": g.get("location"),
            "start": g.get("start", {}),
            "end": g.get("end", {})
        }
        desired_reduced = {
            "summary": body.get("summary"),
            "description": body.get("description"),
            "location": body.get("location"),
            "start": body.get("start"),
            "end": body.get("end")
        }
        if same_enough(desired_reduced, current_reduced):
            skipped += 1
            continue

        gid = g["id"]
        try:
            resp = request_with_backoff(
                lambda: service.events().patch(calendarId=CAL_ID, eventId=gid, body=body, sendUpdates="none"),
                f"events.patch uid={uid}"
            )
            if resp and resp.get("id"):
                updated += 1
            else:
                failed += 1
        except Exception as e:
            print(f"[warn] patch failed uid={uid}: {e}")
            failed += 1

    # 2) DELETE ORPHANS (in window)
    if DELETE_ORPHANS:
        local_keys = set(local.keys())
        for k, g in g_map.items():
            if k in local_keys:
                continue
            gid = g["id"]
            try:
                request_with_backoff(
                    lambda: service.events().delete(calendarId=CAL_ID, eventId=gid, sendUpdates="none"),
                    f"events.delete gid={gid}"
                )
                deleted += 1
            except Exception as e:
                print(f"[warn] delete failed gid={gid}: {e}")
                failed += 1

    print(f"[DONE] Created: {created}, Updated: {updated}, Skipped: {skipped}, Deleted: {deleted}, Failed: {failed}")

if __name__ == "__main__":
    main()
