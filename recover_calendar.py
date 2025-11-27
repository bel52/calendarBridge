#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
CalendarBridge Recovery Script - Repopulates Google Calendar from Outlook ICS
Version 6.0.1 - Fixed for multiple VCALENDAR blocks
"""

import os
import sys
import json
import time
import random
import glob
from datetime import datetime, timedelta, timezone, date
from typing import Dict, Any, List

ROOT = os.path.expanduser("~/calendarBridge")
CONFIG_PATH = os.path.join(ROOT, "calendar_config.json")

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
API_DELAY = 0.15

from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2.credentials import Credentials
from icalendar import Calendar
import recurring_ical_events
from dateutil.tz import gettz

def get_service():
    creds_path = os.path.join(ROOT, "token.json")
    creds = Credentials.from_authorized_user_file(creds_path)
    return build("calendar", "v3", credentials=creds, cache_discovery=False)

def to_iso(dt) -> str:
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
    if isinstance(dt, datetime):
        return dt.date()
    if isinstance(dt, date):
        return dt
    return date(dt.year, dt.month, dt.day)

def is_all_day_event(comp) -> bool:
    ms_allday = comp.get("X-MICROSOFT-CDO-ALLDAYEVENT")
    if ms_allday and str(ms_allday).upper().strip() == "TRUE":
        return True
    dtstart = comp.get("DTSTART")
    if not dtstart:
        return False
    dt_start = dtstart.dt
    if isinstance(dt_start, date) and not isinstance(dt_start, datetime):
        return True
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

def build_event_body(ev: Dict[str, Any], tz: str) -> Dict[str, Any]:
    body = {
        "summary": ev["summary"] or "(No title)",
        "location": ev["location"] or None,
        "description": ev["description"] or None,
        "extendedProperties": {
            "private": {
                "icalUID": ev["uid"],
                "source": "calendarbridge"
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

def parse_ics_with_multiple_calendars(ics_path: str, window_start: datetime, window_end: datetime) -> List[Dict[str, Any]]:
    """Parse ICS file that may contain multiple VCALENDAR blocks."""
    events = []
    all_day_count = 0
    
    with open(ics_path, "rb") as f:
        raw_data = f.read()
    
    # Split by VCALENDAR blocks
    content = raw_data.decode("utf-8", errors="ignore")
    blocks = content.split("BEGIN:VCALENDAR")
    
    for block in blocks:
        if not block.strip():
            continue
        
        # Reconstruct valid VCALENDAR
        ics_block = "BEGIN:VCALENDAR" + block
        if not ics_block.rstrip().endswith("END:VCALENDAR"):
            ics_block = ics_block.rstrip() + "\nEND:VCALENDAR"
        
        try:
            cal = Calendar.from_ical(ics_block.encode("utf-8"))
            
            # Try to expand recurring events
            try:
                expanded = recurring_ical_events.of(cal).between(window_start, window_end)
            except:
                expanded = list(cal.walk("VEVENT"))
            
            for comp in expanded:
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
                all_day = is_all_day_event(comp)
                
                if all_day:
                    all_day_count += 1
                    start_date = normalize_to_date(dtstart)
                    end_date = normalize_to_date(dtend) if dtend else start_date + timedelta(days=1)
                    if end_date <= start_date:
                        end_date = start_date + timedelta(days=1)
                else:
                    start_date = dtstart
                    end_date = dtend if dtend else dtstart + timedelta(hours=1)
                
                events.append({
                    "uid": uid,
                    "summary": (comp.get("SUMMARY") or "").strip(),
                    "location": (comp.get("LOCATION") or "").strip(),
                    "description": (comp.get("DESCRIPTION") or "").strip(),
                    "start": start_date,
                    "end": end_date,
                    "allDay": all_day
                })
        except Exception as e:
            # Skip invalid blocks
            continue
    
    print(f"  Parsed {len(events)} events ({all_day_count} all-day)")
    return events

def main():
    print("=" * 60)
    print("CalendarBridge Recovery Script v6.0.1")
    print("=" * 60)
    print(f"Calendar: {CAL_ID}")
    print(f"Window: -{DAYS_PAST} to +{DAYS_FUTURE} days")
    print()
    
    response = input("This will create events in Google Calendar. Continue? (yes/no): ").strip().lower()
    if response != "yes":
        print("Aborted.")
        return
    
    print()
    print("Loading ICS file...")
    
    ics_path = os.path.join(ROOT, "outbox", "outlook_full_export.ics")
    if not os.path.exists(ics_path):
        print(f"ERROR: ICS file not found at {ics_path}")
        return
    
    tz = gettz(TIMEZONE)
    now = datetime.now(tz)
    window_start = now - timedelta(days=DAYS_PAST)
    window_end = now + timedelta(days=DAYS_FUTURE)
    
    events = parse_ics_with_multiple_calendars(ics_path, window_start, window_end)
    
    if not events:
        print("ERROR: No events parsed")
        return
    
    print()
    print("Connecting to Google Calendar...")
    service = get_service()
    
    print("Checking for existing events...")
    existing_keys = set()
    
    now_utc = datetime.now(timezone.utc)
    time_min = (now_utc - timedelta(days=DAYS_PAST)).isoformat()
    time_max = (now_utc + timedelta(days=DAYS_FUTURE)).isoformat()
    
    page_token = None
    while True:
        resp = service.events().list(
            calendarId=CAL_ID,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            showDeleted=False,
            maxResults=2500,
            pageToken=page_token
        ).execute()
        for item in resp.get("items", []):
            ical_uid = item.get("iCalUID")
            if not ical_uid:
                ext = item.get("extendedProperties", {}).get("private", {}) or {}
                ical_uid = ext.get("icalUID")
            if ical_uid:
                start = item.get("start", {})
                start_str = start.get("date") or start.get("dateTime", "")
                existing_keys.add(f"{ical_uid}|{start_str[:10]}")
        page_token = resp.get("nextPageToken")
        if not page_token:
            break
    
    print(f"Found {len(existing_keys)} existing events in Google")
    print()
    
    print("Creating events...")
    created = 0
    skipped = 0
    failed = 0
    
    for i, ev in enumerate(events):
        start_str = str(ev["start"])[:10]
        key = f"{ev['uid']}|{start_str}"
        
        if key in existing_keys:
            skipped += 1
            continue
        
        body = build_event_body(ev, TIMEZONE)
        time.sleep(API_DELAY + random.uniform(0, 0.05))
        
        try:
            service.events().insert(calendarId=CAL_ID, body=body).execute()
            created += 1
            existing_keys.add(key)  # Prevent duplicates within this run
            
            if created % 100 == 0:
                print(f"  Created {created} events...")
        except HttpError as e:
            if e.resp.status == 409:
                skipped += 1
            elif e.resp.status in (403, 429):
                print(f"  Rate limited, sleeping 30s...")
                time.sleep(30)
                try:
                    service.events().insert(calendarId=CAL_ID, body=body).execute()
                    created += 1
                except:
                    failed += 1
            else:
                failed += 1
        except Exception as e:
            failed += 1
    
    print()
    print("=" * 60)
    print(f"RECOVERY COMPLETE")
    print(f"Created: {created}, Skipped: {skipped}, Failed: {failed}")
    print("=" * 60)
    
    # Clear old state file
    state_path = os.path.join(ROOT, "sync_state.json")
    if os.path.exists(state_path):
        os.rename(state_path, state_path + ".backup")
        print("Old state file backed up")
    
    print()
    print("Next: Run ./full_sync.sh to build state file")

if __name__ == "__main__":
    main()
