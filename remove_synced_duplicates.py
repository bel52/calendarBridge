import os
import datetime
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from collections import defaultdict

# ---- CONFIG ----
SCOPES = ['https://www.googleapis.com/auth/calendar']
CALENDAR_ID = 'primary'
CREDENTIALS_FILE = os.path.expanduser('~/calendarBridge/credentials.json')
TOKEN_FILE = os.path.expanduser('~/calendarBridge/token.json')
LOOKBACK_DAYS = 180
LOOKAHEAD_DAYS = 120

# ---- AUTH ----
creds = None
if os.path.exists(TOKEN_FILE):
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
        creds = flow.run_local_server(port=0)
    with open(TOKEN_FILE, 'w') as token:
        token.write(creds.to_json())

service = build('calendar', 'v3', credentials=creds)

# ---- TIME RANGE ----
now = datetime.datetime.utcnow()
time_min = (now - datetime.timedelta(days=LOOKBACK_DAYS)).isoformat() + 'Z'
time_max = (now + datetime.timedelta(days=LOOKAHEAD_DAYS)).isoformat() + 'Z'

print(f"📅 Fetching events from {time_min} to {time_max}...")

events_result = service.events().list(
    calendarId=CALENDAR_ID,
    timeMin=time_min,
    timeMax=time_max,
    singleEvents=True,
    orderBy='startTime'
).execute()

events = events_result.get('items', [])
print(f"✅ Fetched {len(events)} events.")

# ---- GROUP EVENTS BY SUMMARY+START+END+LOCATION ----
grouped = defaultdict(list)
for event in events:
    key = (
        event.get("summary", "").strip(),
        event.get("start", {}).get("dateTime") or event.get("start", {}).get("date"),
        event.get("end", {}).get("dateTime") or event.get("end", {}).get("date"),
        event.get("location", "").strip()
    )
    grouped[key].append(event)

# ---- FILTER ONLY GROUPS WITH SCRIPT TAGGED UID ----
groups_to_deduplicate = {
    k: v for k, v in grouped.items()
    if len(v) > 1 and all(
        e.get("extendedProperties", {}).get("private", {}).get("icalUID")
        for e in v
    )
}

# ---- DELETE DUPLICATES (KEEP FIRST CREATED) ----
deleted_count = 0
for group, dupes in groups_to_deduplicate.items():
    dupes_sorted = sorted(dupes, key=lambda e: e.get('created'))
    to_delete = dupes_sorted[1:]
    for event in to_delete:
        try:
            service.events().delete(calendarId=CALENDAR_ID, eventId=event['id']).execute()
            deleted_count += 1
            print(f"🗑 Deleted duplicate: {group[0]} @ {event['start']}")
        except Exception as e:
            print(f"⚠️ Failed to delete: {group[0]} @ {event['start']} - {e}")

print(f"\n✅ Done. Total duplicates deleted: {deleted_count}")
