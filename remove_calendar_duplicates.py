#!/usr/bin/env python3

from collections import defaultdict
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
import datetime

# Authenticate to Google Calendar
creds = Credentials.from_authorized_user_file("token.json", ["https://www.googleapis.com/auth/calendar"])
service = build("calendar", "v3", credentials=creds)

now = datetime.datetime.utcnow().isoformat() + 'Z'
print("📅 Fetching events from Google Calendar...")

events_result = service.events().list(
    calendarId='primary',
    timeMin=now,
    maxResults=2500,
    singleEvents=True
).execute()

events = events_result.get('items', [])
print(f"✅ Fetched {len(events)} events.")

seen = defaultdict(list)
for event in events:
    key = (
        event.get('summary', '').strip(),
        event.get('start', {}).get('dateTime') or event.get('start', {}).get('date'),
        event.get('location', '').strip()
    )
    seen[key].append(event)

deleted = 0
for key, dupes in seen.items():
    if len(dupes) > 1:
        # Keep the first one, delete the rest
        for event in dupes[1:]:
            try:
                service.events().delete(calendarId='primary', eventId=event['id']).execute()
                print(f"🗑 Deleted: {event['summary']} @ {event['start']}")
                deleted += 1
            except Exception as e:
                print(f"❌ Failed to delete {event['summary']}: {e}")

print(f"\n🧹 Cleanup complete. {deleted} duplicates removed.")
