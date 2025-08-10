import os
import json
import glob
from datetime import datetime, timedelta
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials
from icalendar import Calendar

GOOGLE_CALENDAR_ID = 'primary'
ICS_DIR = os.path.expanduser('~/calendarBridge/outbox/')
SYNC_TRACKER_FILE = 'synced_events.json'

def load_google_credentials():
    creds = Credentials.from_authorized_user_file('token.json', ['https://www.googleapis.com/auth/calendar'])
    return build('calendar', 'v3', credentials=creds)

def load_synced_events():
    if os.path.exists(SYNC_TRACKER_FILE):
        with open(SYNC_TRACKER_FILE, 'r') as f:
            return json.load(f)
    return {}

def save_synced_events(data):
    with open(SYNC_TRACKER_FILE, 'w') as f:
        json.dump(data, f, indent=2)

def parse_ics_events():
    events = {}
    for ics_file in glob.glob(os.path.join(ICS_DIR, '*.ics')):
        with open(ics_file, 'rb') as f:
            cal = Calendar.from_ical(f.read())
            for component in cal.walk():
                if component.name == "VEVENT":
                    summary = str(component.get('summary'))
                    start = component.get('dtstart').dt
                    end = component.get('dtend').dt
                    uid = str(component.get('uid'))
                    events[uid] = {
                        'summary': summary,
                        'start': start.isoformat(),
                        'end': end.isoformat()
                    }
    return events

def fetch_google_events(service):
    now = datetime.utcnow().isoformat() + 'Z'
    time_max = (datetime.utcnow() + timedelta(days=120)).isoformat() + 'Z'
    events = {}
    page_token = None
    while True:
        response = service.events().list(calendarId=GOOGLE_CALENDAR_ID, timeMin=now, timeMax=time_max,
                                         singleEvents=True, orderBy='startTime',
                                         pageToken=page_token).execute()
        for item in response.get('items', []):
            uid = item.get('extendedProperties', {}).get('private', {}).get('icalUID')
            if uid:
                events[uid] = item
        page_token = response.get('nextPageToken')
        if not page_token:
            break
    return events

def main():
    service = load_google_credentials()
    outlook_events = parse_ics_events()
    synced = load_synced_events()
    google_events = fetch_google_events(service)

    to_add = {}
    to_delete = {}

    for uid, event in outlook_events.items():
        if uid not in synced:
            to_add[uid] = event

    for uid in synced:
        if uid not in outlook_events:
            to_delete[uid] = synced[uid]

    # Delete removed events from Google
    for uid, event in to_delete.items():
        try:
            service.events().delete(calendarId=GOOGLE_CALENDAR_ID, eventId=event['id']).execute()
            print(f"🗑 Deleted Google event: {event['summary']} ({event['start']})")
            del synced[uid]
        except Exception as e:
            print(f"⚠️ Could not delete {uid}: {e}")

    # Add new events to Google
    for uid, event in to_add.items():
        body = {
            'summary': event['summary'],
            'start': {'dateTime': event['start'], 'timeZone': 'America/New_York'},
            'end': {'dateTime': event['end'], 'timeZone': 'America/New_York'},
            'extendedProperties': {'private': {'icalUID': uid}}
        }
        try:
            created = service.events().insert(calendarId=GOOGLE_CALENDAR_ID, body=body).execute()
            print(f"✅ Added Google event: {event['summary']} ({event['start']})")
            synced[uid] = {'id': created['id'], 'summary': event['summary'], 'start': event['start']}
        except Exception as e:
            print(f"⚠️ Could not add event {uid}: {e}")

    save_synced_events(synced)

if __name__ == '__main__':
    main()
