#!/usr/bin/env python3

import hashlib
import datetime
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

def generate_icaluid(summary, start):
    uid_seed = f"{summary.strip()}-{start.strip()}"
    return hashlib.sha256(uid_seed.encode()).hexdigest()[:32]

# Setup
creds = Credentials.from_authorized_user_file("token.json", ["https://www.googleapis.com/auth/calendar"])
service = build("calendar", "v3", credentials=creds)

# Time window: 6 months ago to 120 days in the future
start_window = (datetime.datetime.utcnow() - datetime.timedelta(days=180)).isoformat() + 'Z'
end_window = (datetime.datetime.utcnow() + datetime.timedelta(days=120)).isoformat() + 'Z'

print(f"🔍 Scanning events between {start_window} and {end_window}...")

page_token = None
updated = 0

while True:
    events_result = service.events().list(
        calendarId='primary',
        timeMin=start_window,
        timeMax=end_window,
        singleEvents=True,
        showDeleted=False,
        maxResults=2500,
        pageToken=page_token
    ).execute()

    for event in events_result.get('items', []):
        extended = event.get("extendedProperties", {}).get("private", {})
        if "icalUID" in extended:
            continue  # Already tagged

        summary = event.get("summary", "").strip()
        start = event.get("start", {}).get("dateTime") or event.get("start", {}).get("date")
        if not summary or not start:
            continue

        uid = generate_icaluid(summary, start)
        event.setdefault("extendedProperties", {}).setdefault("private", {})["icalUID"] = uid

        try:
            service.events().update(calendarId='primary', eventId=event['id'], body=event).execute()
            print(f"✅ Tagged: {summary} @ {start} with UID {uid}")
            updated += 1
        except Exception as e:
            print(f"❌ Failed to update {summary} @ {start}: {e}")

    page_token = events_result.get('nextPageToken')
    if not page_token:
        break

print(f"\n✅ Tagging complete. {updated} events updated.")
