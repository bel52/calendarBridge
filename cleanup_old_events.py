#!/usr/bin/env python3
import json
import os
import random
import time
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

def delete_event_with_backoff(service, calendar_id, event_id):
    attempts = 0
    while attempts < 5:
        try:
            service.events().delete(calendarId=calendar_id, eventId=event_id).execute()
            print(f"Deleted old event {event_id}")
            return
        except HttpError as e:
            status = getattr(e.resp, 'status', None)
            if status in (403, 429):
                # Rate limit – back off and retry
                attempts += 1
                sleep_s = min(2.0 * attempts, 6.0) + random.uniform(0.1, 0.5)
                print(f"Rate limit for {event_id}, sleeping {sleep_s:.1f}s (attempt {attempts}/5)")
                time.sleep(sleep_s)
                continue
            if status == 410:
                print(f"{event_id} already deleted")
                return
            print(f"Failed to delete {event_id}: {e}")
            return

def main():
    with open('calendar_config.json', 'r') as f:
        calendar_id = json.load(f).get('google_calendar_id', 'primary')

    old_state_file = 'sync_state_backup.json'
    if not os.path.exists(old_state_file):
        print(f"Backup state file '{old_state_file}' not found. Aborting.")
        return

    with open(old_state_file, 'r') as f:
        old_state = json.load(f)

    print(f"Attempting to delete {len(old_state)} old events...")

    scopes = ['https://www.googleapis.com/auth/calendar']
    creds = Credentials.from_authorized_user_file('token.json', scopes)
    service = build('calendar', 'v3', credentials=creds)

    for event_id in old_state.values():
        delete_event_with_backoff(service, calendar_id, event_id)

    print("Cleanup complete.")

if __name__ == '__main__':
    main()
