#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Cleanup duplicate Google Calendar events.

Default behavior:
- Scans the same date window used by safe_sync.py (config-driven).
- Groups events by (UID, normalized start, normalized end, summary).
- Keeps ONE event per group (oldest by created time).
- Deletes only events tagged as ours:
    extendedProperties.private.source == "CalendarBridge"

Optional aggressive mode:
- If --include-all is passed, deletes ALL extra copies in each group
  (regardless of source) and keeps just one.

Usage:
    ./cleanup_duplicates.py                      # dry-run, conservative
    ./cleanup_duplicates.py --apply              # delete safe (CalendarBridge) duplicates
    ./cleanup_duplicates.py --apply --include-all  # delete all extras per group
"""

import os
import sys
import json
import time
import logging
from datetime import datetime, timedelta, timezone
from typing import Dict, Any, Tuple, List, Optional

from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

ROOT = os.path.dirname(os.path.abspath(__file__))
CONFIG_PATH = os.path.join(ROOT, "calendar_config.json")
TOKEN_PATH = os.path.join(ROOT, "token.json")
CREDENTIALS_PATH = os.path.join(ROOT, "credentials.json")
LOG_DIR = os.path.join(ROOT, "logs")
os.makedirs(LOG_DIR, exist_ok=True)
LOG_PATH = os.path.join(LOG_DIR, "cleanup_duplicates.log")
REPORT_PATH = os.path.join(LOG_DIR, "cleanup_duplicates_report.json")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_PATH),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("calendarbridge.cleanup")

ORPHAN_MARKER = "CalendarBridge"


# ------------- Config & Auth -------------

def load_config() -> Dict[str, Any]:
    if not os.path.exists(CONFIG_PATH):
        raise SystemExit(f"Missing config file: {CONFIG_PATH}")
    with open(CONFIG_PATH, "r") as f:
        cfg = json.load(f)
    required = ["google_calendar_id", "sync_days_past", "sync_days_future"]
    missing = [k for k in required if k not in cfg]
    if missing:
        raise SystemExit(f"Missing keys in calendar_config.json: {missing}")
    return cfg


def get_google_service(scopes=None):
    if scopes is None:
        scopes = ["https://www.googleapis.com/auth/calendar"]
    creds = None
    if os.path.exists(TOKEN_PATH):
        try:
            creds = Credentials.from_authorized_user_file(TOKEN_PATH, scopes)
        except Exception as e:
            log.warning(f"Failed to load token.json, re-auth required: {e}")
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

    return build("calendar", "v3", credentials=creds)


def get_time_window_iso(cfg: Dict[str, Any]) -> Tuple[str, str]:
    now = datetime.now(timezone.utc)
    time_min = now - timedelta(days=int(cfg["sync_days_past"]))
    time_max = now + timedelta(days=int(cfg["sync_days_future"]))
    return time_min.isoformat(), time_max.isoformat()


# ------------- Helpers -------------

def _normalize_start(start: Dict[str, Any]) -> Optional[str]:
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


def _normalize_end(end: Dict[str, Any]) -> Optional[str]:
    if not end:
        return None
    if "date" in end:
        return end["date"]
    dt_str = end.get("dateTime")
    if not dt_str:
        return None
    if "T" not in dt_str:
        return dt_str
    date_part, time_part = dt_str.split("T", 1)
    time_part = time_part.split("+")[0].split("-")[0]
    time_part = time_part[:8]
    return f"{date_part}T{time_part}"


def is_our_event(ev: Dict[str, Any]) -> bool:
    ext = (ev.get("extendedProperties") or {}).get("private") or {}
    return ext.get("source") == ORPHAN_MARKER


def get_uid(ev: Dict[str, Any]) -> Optional[str]:
    ext = (ev.get("extendedProperties") or {}).get("private") or {}
    return ext.get("icalUID") or ev.get("iCalUID")


# ------------- Fetch & Group -------------

def fetch_events(service, calendar_id: str, cfg: Dict[str, Any]) -> List[Dict[str, Any]]:
    time_min, time_max = get_time_window_iso(cfg)
    events: List[Dict[str, Any]] = []
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
                    orderBy="startTime",
                    pageToken=page_token,
                )
                .execute()
            )
        except HttpError as e:
            log.error(f"Failed to fetch events: {e}")
            break

        items = resp.get("items", [])
        total += len(items)
        events.extend(items)

        page_token = resp.get("nextPageToken")
        if not page_token:
            break

    log.info(f"Fetched {total} events in cleanup window")
    return events


def group_duplicates(events: List[Dict[str, Any]]) -> Dict[Tuple[str, str, str, str], List[Dict[str, Any]]]:
    """
    Group events by (uid, normalized_start, normalized_end, summary)
    Only groups with len > 1 are considered duplicates.
    """
    groups: Dict[Tuple[str, str, str, str], List[Dict[str, Any]]] = {}

    for ev in events:
        uid = get_uid(ev)
        if not uid:
            continue

        start_key = _normalize_start(ev.get("start") or {})
        end_key = _normalize_end(ev.get("end") or {})
        if not start_key or not end_key:
            continue

        summary = (ev.get("summary") or "").strip()
        key = (uid, start_key, end_key, summary)

        groups.setdefault(key, []).append(ev)

    dup_groups = {k: v for k, v in groups.items() if len(v) > 1}
    log.info(f"Identified {len(dup_groups)} duplicate groups (>=2 events per group)")
    return dup_groups


def pick_keep_and_delete(events: List[Dict[str, Any]]) -> Tuple[Dict[str, Any], List[Dict[str, Any]]]:
    """
    Given a list of duplicate events, choose one to keep (oldest by created time),
    and return (keep_event, [events_to_delete]).
    """
    def created_ts(ev: Dict[str, Any]) -> float:
        created_str = ev.get("created") or ev.get("updated")
        if not created_str:
            return 0.0
        try:
            return datetime.fromisoformat(created_str.replace("Z", "+00:00")).timestamp()
        except Exception:
            return 0.0

    sorted_events = sorted(events, key=created_ts)
    keep = sorted_events[0]
    delete = sorted_events[1:]
    return keep, delete


# ------------- Main cleanup -------------

def main():
    apply = "--apply" in sys.argv
    include_all = "--include-all" in sys.argv

    cfg = load_config()
    cal_id = cfg["google_calendar_id"]
    service = get_google_service()

    log.info("=" * 60)
    log.info(f"Duplicate cleanup starting for calendar {cal_id}")
    log.info(f"Apply mode: {'YES (will delete)' if apply else 'NO (dry-run only)'}")
    log.info(f"Include-all mode: {'YES (delete all extras per group)' if include_all else 'NO (CalendarBridge events only)'}")

    events = fetch_events(service, cal_id, cfg)
    dup_groups = group_duplicates(events)

    report = {
        "apply": apply,
        "include_all": include_all,
        "total_groups": len(dup_groups),
        "groups": [],
    }

    total_safe = 0
    total_deleted = 0

    for key, group in dup_groups.items():
        uid, start_key, end_key, summary = key
        keep, candidates = pick_keep_and_delete(group)

        if include_all:
            safe_deletes = list(candidates)  # delete all extras
            unsafe = []
        else:
            safe_deletes = [ev for ev in candidates if is_our_event(ev)]
            unsafe = [ev for ev in candidates if not is_our_event(ev)]

        group_entry = {
            "uid": uid,
            "summary": summary,
            "start": start_key,
            "end": end_key,
            "keep_id": keep.get("id"),
            "keep_created": keep.get("created"),
            "delete_ids_safe": [ev.get("id") for ev in safe_deletes],
            "delete_ids_unsafe": [ev.get("id") for ev in unsafe],
        }
        report["groups"].append(group_entry)
        total_safe += len(safe_deletes)

        if apply and safe_deletes:
            for ev in safe_deletes:
                gid = ev.get("id")
                if not gid:
                    continue
                try:
                    log.info(f"Deleting duplicate event {gid} (summary='{summary}', uid={uid})")
                    service.events().delete(calendarId=cal_id, eventId=gid).execute()
                    total_deleted += 1
                    time.sleep(0.05)
                except HttpError as e:
                    log.error(f"Failed to delete {gid}: {e}")

    with open(REPORT_PATH, "w") as f:
        json.dump(report, f, indent=2)

    log.info("=" * 60)
    log.info(f"Duplicate cleanup complete. Safe duplicates identified: {total_safe}")
    if apply:
        log.info(f"Deleted: {total_deleted} events")
    else:
        log.info("Dry run only; no events were deleted.")
        log.info(f"Review report at {REPORT_PATH} and re-run with --apply (and optionally --include-all) if satisfied.")
    log.info("=" * 60)


if __name__ == "__main__":
    main()
