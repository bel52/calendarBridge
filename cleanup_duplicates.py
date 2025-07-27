import json
from datetime import datetime, timedelta
from googleapiclient.discovery import build
from google.oauth2.credentials import Credentials

GOOGLE_CALENDAR_ID = 'primary'

def load_google_credentials():
    creds = Credentials.from_authorized_user_file('token.json', ['https://www.googleapis.com/auth/calendar'])
    return build('calendar', 'v3', credentials=creds)

def get_time_range():
    now = datetime.utcnow()
    past = (now - timedelta(days=180)).isoformat() + 'Z'
    future = (now + timedelta(days=120)).isoformat() + 'Z'
    return past, future

def fetch_events(service, start_time, end_time):
    all_events = []
    page_token = None
    while True:
        response = service.events().list(calendarId=GOOGLE_CALENDAR_ID,
                                         timeMin=start_time,
                                         timeMax=end_time,
                                         singleEvents=True,
                                         orderBy='startTime',
                                         pageToken=page_token).execute()
        all_events.extend(response.get('items', []))
        page_token = response.get('nextPageToken')
        if not page_token:
            break
    return all_events

def main():
    service = load_google_credentials()
    start_time, end_time = get_time_range()
    events = fetch_events(service, start_time, end_time)

    seen = {}
    deleted = 0

    for event in events:
        if event.get('recurringEventId'):
            continue  # skip recurring
        key = (event['summary'], event['start'].get('dateTime') or event['start'].get('date'))
        if key in seen:
            try:
                service.events().delete(calendarId=GOOGLE_CALENDAR_ID, eventId=event['id']).execute()
                print(f"🗑 Deleted duplicate: {key}")
                deleted += 1
            except Exception as e:
                print(f"⚠️ Failed to delete: {e}")
        else:
            seen[key] = event['id']

    print(f"\n✅ Cleanup complete. {deleted} duplicates removed.")

if __name__ == '__main__':
    main()
