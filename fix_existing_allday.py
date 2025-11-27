#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
fix_existing_allday.py - One-time fix for all-day events in Google Calendar

This script finds all events that SHOULD be all-day (based on midnight-to-midnight
times or our source data) and converts them to proper all-day format for iOS.

Run once after updating safe_sync.py, then let the regular sync maintain things.
"""

import os
import sys
import time
import random
from datetime import datetime, timedelta, timezone
from typing import Dict, Any, List

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Configuration
ROOT = os.path.expanduser("~/calendarBridge")
CAL_ID = "brett@leathermans.net"  # Your calendar
DRY_RUN = False  # Set to True to preview changes without applying

def get_service():
    creds_path = os.path.join(ROOT, "token.json")
    creds = Credentials.from_authorized_user_file(creds_path)
    return build("calendar", "v3", credentials=creds, cache_discovery=False)

def sleep_with_backoff(attempt: int):
    """Sleep with exponential backoff."""
    wait = min(2 ** attempt + random.uniform(0, 1), 60)
    print(f"    Rate limited, sleeping {wait:.1f}s...")
    time.sleep(wait)

def is_midnight_to_midnight(event: Dict[str, Any]) -> bool:
    """Check if a timed event spans midnight-to-midnight (should be all-day)."""
    start = event.get("start", {})
    end = event.get("end", {})
    
    # Already all-day format
    if "date" in start and "dateTime" not in start:
        return False  # Already correct
    
    start_dt = start.get("dateTime", "")
    end_dt = end.get("dateTime", "")
    
    if not start_dt or not end_dt:
        return False
    
    # Check for T00:00:00 pattern (midnight)
    # Could be T00:00:00-05:00, T00:00:00Z, etc.
    start_midnight = "T00:00:00" in start_dt
    end_midnight = "T00:00:00" in end_dt
    
    if start_midnight and end_midnight:
        # Parse to verify it's actually full days
        try:
            # Extract just the date parts
            start_date = start_dt.split("T")[0]
            end_date = end_dt.split("T")[0]
            
            from datetime import date
            sd = date.fromisoformat(start_date)
            ed = date.fromisoformat(end_date)
            
            # Must span at least one full day
            return ed > sd
        except Exception:
            pass
    
    return False

def fix_event(service, event: Dict[str, Any]) -> bool:
    """Convert a midnight-to-midnight event to proper all-day format."""
    event_id = event.get("id")
    if not event_id:
        return False
    
    start = event.get("start", {})
    end = event.get("end", {})
    
    start_dt = start.get("dateTime", "")
    end_dt = end.get("dateTime", "")
    
    # Extract dates
    start_date = start_dt.split("T")[0]
    end_date = end_dt.split("T")[0]
    
    # Build patch body - convert to all-day format
    patch_body = {
        "start": {"date": start_date},
        "end": {"date": end_date},
        "transparency": "transparent"
    }
    
    if DRY_RUN:
        print(f"    [DRY RUN] Would patch: {event.get('summary', 'No title')}")
        print(f"              From: {start_dt} → {end_dt}")
        print(f"              To:   {start_date} → {end_date} (all-day, transparent)")
        return True
    
    # Apply patch with retry
    for attempt in range(5):
        try:
            time.sleep(0.5 + random.uniform(0, 0.3))  # Rate limiting
            service.events().patch(
                calendarId=CAL_ID,
                eventId=event_id,
                body=patch_body,
                sendUpdates="none"
            ).execute()
            return True
        except HttpError as e:
            if e.resp.status in (403, 429) and attempt < 4:
                sleep_with_backoff(attempt)
                continue
            print(f"    [ERROR] Failed to patch {event_id}: {e}")
            return False
        except Exception as e:
            print(f"    [ERROR] Unexpected error for {event_id}: {e}")
            return False
    
    return False

def main():
    print("=" * 60)
    print("Fix Existing All-Day Events")
    print("=" * 60)
    
    if DRY_RUN:
        print("[MODE] DRY RUN - No changes will be made")
    else:
        print("[MODE] LIVE - Changes will be applied")
    print()
    
    service = get_service()
    
    # Get events from past 60 days to future 90 days
    now = datetime.now(timezone.utc)
    time_min = (now - timedelta(days=60)).isoformat()
    time_max = (now + timedelta(days=90)).isoformat()
    
    print(f"[INFO] Scanning events from {time_min[:10]} to {time_max[:10]}")
    print()
    
    # Fetch all events
    all_events: List[Dict[str, Any]] = []
    page_token = None
    
    while True:
        try:
            resp = service.events().list(
                calendarId=CAL_ID,
                timeMin=time_min,
                timeMax=time_max,
                singleEvents=True,
                showDeleted=False,
                maxResults=2500,
                pageToken=page_token,
                orderBy="startTime"
            ).execute()
            
            all_events.extend(resp.get("items", []))
            page_token = resp.get("nextPageToken")
            if not page_token:
                break
        except Exception as e:
            print(f"[ERROR] Failed to list events: {e}")
            sys.exit(1)
    
    print(f"[INFO] Found {len(all_events)} total events")
    
    # Find events needing fix
    needs_fix = []
    already_allday = 0
    timed_events = 0
    
    for event in all_events:
        start = event.get("start", {})
        
        # Check if already proper all-day
        if "date" in start and "dateTime" not in start:
            # Check transparency
            if event.get("transparency") != "transparent":
                needs_fix.append(("transparency", event))
            else:
                already_allday += 1
            continue
        
        # Check if it's a midnight-to-midnight event
        if is_midnight_to_midnight(event):
            needs_fix.append(("convert", event))
        else:
            timed_events += 1
    
    print(f"[INFO] Already correct all-day: {already_allday}")
    print(f"[INFO] Timed events (no fix needed): {timed_events}")
    print(f"[INFO] Events needing fix: {len(needs_fix)}")
    print()
    
    if not needs_fix:
        print("[DONE] No events need fixing!")
        return
    
    # Show what needs fixing
    print("Events to fix:")
    print("-" * 60)
    
    convert_count = 0
    transparency_count = 0
    
    for fix_type, event in needs_fix[:20]:  # Show first 20
        summary = event.get("summary", "No title")[:40]
        start = event.get("start", {})
        
        if fix_type == "convert":
            print(f"  [CONVERT]     {summary}")
            print(f"                {start.get('dateTime', '')[:16]}")
            convert_count += 1
        else:
            print(f"  [TRANSPARENCY] {summary}")
            transparency_count += 1
    
    if len(needs_fix) > 20:
        print(f"  ... and {len(needs_fix) - 20} more")
    
    print()
    print(f"Summary: {convert_count} to convert, {transparency_count} transparency fixes")
    print()
    
    if DRY_RUN:
        print("[DRY RUN] Set DRY_RUN = False to apply changes")
        return
    
    # Confirm
    response = input("Apply fixes? (yes/no): ").strip().lower()
    if response != "yes":
        print("Aborted.")
        return
    
    # Apply fixes
    print()
    print("Applying fixes...")
    fixed = 0
    failed = 0
    
    for i, (fix_type, event) in enumerate(needs_fix):
        summary = event.get("summary", "No title")[:30]
        
        if fix_type == "convert":
            print(f"[{i+1}/{len(needs_fix)}] Converting: {summary}")
            if fix_event(service, event):
                fixed += 1
            else:
                failed += 1
        else:
            # Just fix transparency
            event_id = event.get("id")
            if event_id:
                try:
                    time.sleep(0.3)
                    service.events().patch(
                        calendarId=CAL_ID,
                        eventId=event_id,
                        body={"transparency": "transparent"},
                        sendUpdates="none"
                    ).execute()
                    print(f"[{i+1}/{len(needs_fix)}] Fixed transparency: {summary}")
                    fixed += 1
                except Exception as e:
                    print(f"[{i+1}/{len(needs_fix)}] Failed: {summary} - {e}")
                    failed += 1
    
    print()
    print("=" * 60)
    print(f"[DONE] Fixed: {fixed}, Failed: {failed}")
    print("=" * 60)

if __name__ == "__main__":
    main()
