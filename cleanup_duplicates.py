#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
cleanup_duplicates.py

Cleanup script to remove duplicate events from Google Calendar
for keys (iCalUID|normalized start), keeping exactly ONE event
per key.

Key points:
- Groups ALL events (ours + non-ours) by the same key safe_sync.py uses:
    key = f"{iCalUID}|<normalized start>"
- For each group with >= 2 events:
    * Choose ONE canonical event to KEEP:
        - Prefer event whose ID matches sync_state["google_ids"][key].
        - Else, prefer one with extendedProperties.private.source == "calendarbridge".
        - Else, keep the first.
    * DELETE all other events in the group.
- We NEVER delete all events for a key; at least one survives.

Default mode:
- DRY RUN (no deletions). Use --apply to actually delete.
"""

import os
import sys
import json
import time
import random
import argparse
from datetime import datetime, timedelta, timezone
from typing import Dict, Any, List, Tuple

ROOT = os.path.expanduser("~/calendarBridge")
CONFIG_PATH = os.path.join(ROOT, "calendar_config.json")
STATE_PATH = os.path.join(ROOT, "sync_state.json")

LOG_PREFIX = "[CLEANUP]"

# --------- Config helpers ---------

def load_config(path: str) -> Dict[str, Any]:
    if os.path.exists(path):
        with open(path, "r") as f:
            return json.load(f)
    return {}

CONF = load_config(CONFIG_PATH)

TIMEZONE = CONF.get("timezone", "America/New_York")
CAL_ID = CONF.get("google_calendar_id", "primary")
DAYS_PAST = int(CONF.get("sync_days_past", 60))
DAYS_FUTURE = int(CONF.get("sync_days_future", 90))
API_DELAY = float(CONF.get("api_delay_seconds", 0.15))
QUOTA_USER = os.environ.get("CALBRIDGE_QUOTA_USER")
ORPHAN_MARKER = "calendarbridge"

# --------- Google API + backoff ---------

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials

def get_google_service():
    creds_path = os.path.join(ROOT, "token.json")
    creds = Credentials.from_authorized_user_file(creds_path)
    return build("calendar", "v3", credentials=creds, cache_discovery=False)

def sleep_with_jitter(base_seconds: float) -> None:
    if base_seconds <= 0:
        return
    jitter = base_seconds * random.uniform(-0.20, 0.20)
    time.sleep(max(0.0, base_seconds + jitter))

def is_rate_limit_error(e: HttpError) -> bool:
    try:
        status = e.resp.status if hasattr(e, "resp") else None
        if status in (429,):
            return True
        if status == 403:
            content = getattr(e, "content", b"") or b""
            msg = content.decode("utf-8", errors="ignore")
            return "rateLimitExceeded" in msg or "userRateLimitExceeded" in msg
    except Exception:
        pass
    return False

def request_with_backoff(service, req_builder, op_desc: str, max_attempts: int = 5) -> Any:
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
                print(f"{LOG_PREFIX} rate-limited on {op_desc}, sleeping {wait:.1f}s (attempt {attempt})")
                time.sleep(wait)
                backoff = min(backoff * 2.0, 30.0)
                continue
            raise
        except Exception as e:
            if attempt < max_attempts:
                wait = backoff + random.uniform(0.0, 0.5)
                print(f"{LOG_PREFIX} error on {op_desc}: {e}, retrying in {wait:.1f}s")
                time.sleep(wait)
                backoff = min(backoff * 2.0, 30.0)
                continue
            raise

# --------- State helpers (read-only) ---------

def load_state(path: str) -> Dict[str, Any]:
    if not os.path.exists(path):
        return {"events": {}, "google_ids": {}}
    try:
        with open(path, "r") as f:
            return json.load(f)
    except Exception:
        return {"events": {}, "google_ids": {}}

# --------- Key & ownership helpers (MUST MATCH safe_sync) ---------

def normalize_start_string(start_obj: Dict[str, Any]) -> str:
    """
    Convert a Google event 'start' object into the same normalized
    string safe_sync uses: YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS
    with any timezone/offset stripped.
    """
    start_str = start_obj.get("date") or start_obj.get("dateTime", "")
    if not start_str:
        return ""
    if "T" not in start_str:
        # all-day
        return start_str
    parts = start_str.split("T")
    rest = parts[1]

    # Remove timezone (e.g. +HH:MM, -HH:MM, or 'Z')
    if "+" in rest:
        rest = rest.split("+", 1)[0]
    elif "-" in rest[1:]:
        rest = rest.split("-", 1)[0]
    elif rest.endswith("Z"):
        rest = rest[:-1]

    time_part = rest[:8]
    return f"{parts[0]}T{time_part}"

def make_key_from_event(item: Dict[str, Any]) -> str | None:
    ical_uid = item.get("iCalUID")
    if not ical_uid:
        ext = item.get("extendedProperties", {}).get("private", {}) or {}
        ical_uid = ext.get("icalUID")
    if not ical_uid:
        return None
    start = item.get("start", {})
    start_str = normalize_start_string(start)
    if not start_str:
        return None
    return f"{ical_uid}|{start_str}"

def is_our_event(ev: Dict[str, Any]) -> bool:
    ext = ev.get("extendedProperties", {}).get("private", {}) or {}
    return ext.get("source") == ORPHAN_MARKER

# --------- Fetch events ---------

def get_time_window_iso() -> Tuple[str, str]:
    now = datetime.now(timezone.utc)
    time_min = now - timedelta(days=DAYS_PAST)
    time_max = now + timedelta(days=DAYS_FUTURE)
    return time_min.isoformat(), time_max.isoformat()

def fetch_all_events(service) -> List[Dict[str, Any]]:
    time_min, time_max = get_time_window_iso()
    print(f"{LOG_PREFIX} scanning window {time_min} → {time_max}")
    events: List[Dict[str, Any]] = []
    page_token = None
    while True:
        resp = request_with_backoff(
            service,
            lambda: service.events().list(
                calendarId=CAL_ID,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                showDeleted=False,
                maxResults=2500,
                pageToken=page_token,
                orderBy="startTime",
            ),
            "events.list",
        )
        events.extend(resp.get("items", []))
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    print(f"{LOG_PREFIX} fetched {len(events)} events in window")
    return events

# --------- Analyze duplicate groups ---------

def analyze_duplicates(events: List[Dict[str, Any]], state: Dict[str, Any]):
    """
    Analyze duplicates per key.

    We group ALL events by key, then look for keys where there are
    2 or more events. For each such group we will keep exactly one.
    """
    groups: Dict[str, List[Dict[str, Any]]] = {}

    for ev in events:
        key = make_key_from_event(ev)
        if not key:
            continue
        groups.setdefault(key, []).append(ev)

    state_google_ids: Dict[str, str] = state.get("google_ids", {})

    dup_info: Dict[str, Dict[str, Any]] = {}

    for key, evs in groups.items():
        if len(evs) <= 1:
            continue

        # Build ID lists
        ids = [e.get("id") for e in evs if e.get("id")]
        if not ids:
            continue

        # Choose canonical to KEEP
        canonical_state_id = state_google_ids.get(key)
        keep_id = None

        # 1) Prefer the one referenced in sync_state.json
        if canonical_state_id and canonical_state_id in ids:
            keep_id = canonical_state_id
        else:
            # 2) Prefer one marked as our event
            ours = [e for e in evs if is_our_event(e) and e.get("id")]
            if ours:
                keep_id = ours[0].get("id")
            else:
                # 3) Fallback to first
                keep_id = ids[0]

        delete_ids = [i for i in ids if i != keep_id]

        ours_ids = [e.get("id") for e in evs if is_our_event(e) and e.get("id")]
        not_ours_ids = [e.get("id") for e in evs if (not is_our_event(e)) and e.get("id")]

        dup_info[key] = {
            "count_total": len(evs),
            "keep_id": keep_id,
            "delete_ids": delete_ids,
            "ours_ids": ours_ids,
            "not_ours_ids": not_ours_ids,
            "sample_summary": evs[0].get("summary", ""),
            "start": evs[0].get("start", {}),
        }

    return dup_info

# --------- Deletion ---------

def delete_duplicates(service, dup_info: Dict[str, Dict[str, Any]], apply: bool):
    total_groups = len(dup_info)
    total_delete = sum(len(v["delete_ids"]) for v in dup_info.values())
    print(f"{LOG_PREFIX} found {total_groups} duplicate groups (keys with >=2 events); {total_delete} events marked for deletion.")

    if not dup_info:
        return

    print(f"{LOG_PREFIX} sample groups:")
    shown = 0
    for key, info in dup_info.items():
        print(f"  key={key}")
        print(f"    summary:        {info['sample_summary']}")
        print(f"    total events:   {info['count_total']}")
        print(f"    ours:           {info['ours_ids']}")
        print(f"    not ours:       {info['not_ours_ids']}")
        print(f"    KEEP:           {info['keep_id']}")
        print(f"    DELETE:         {info['delete_ids']}")
        shown += 1
        if shown >= 5:
            break

    if not apply:
        print(f"{LOG_PREFIX} DRY RUN ONLY – no deletions performed. Re-run with --apply to actually delete.")
        return

    print(f"{LOG_PREFIX} APPLY MODE – deleting {total_delete} events...")
    deleted = 0
    failed = 0
    for key, info in dup_info.items():
        for gid in info["delete_ids"]:
            try:
                request_with_backoff(
                    service,
                    lambda gid=gid: service.events().delete(
                        calendarId=CAL_ID,
                        eventId=gid,
                        sendUpdates="none",
                    ),
                    f"delete duplicate {gid}",
                )
                deleted += 1
            except HttpError as e:
                if e.resp.status == 410:
                    deleted += 1
                else:
                    failed += 1
                    print(f"{LOG_PREFIX} delete failed for {gid}: {e.resp.status}")
            except Exception as e:
                failed += 1
                print(f"{LOG_PREFIX} delete error for {gid}: {e}")

    print(f"{LOG_PREFIX} deletion complete. Deleted={deleted}, Failed={failed}")

# --------- Main ---------

def main():
    parser = argparse.ArgumentParser(
        description="Cleanup duplicate events from Google Calendar (keep exactly one per iCalUID/start key)."
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Actually delete duplicates (default is dry-run).",
    )
    args = parser.parse_args()

    state = load_state(STATE_PATH)
    service = get_google_service()

    events = fetch_all_events(service)
    dup_info = analyze_duplicates(events, state)
    delete_duplicates(service, dup_info, apply=args.apply)

if __name__ == "__main__":
    main()
