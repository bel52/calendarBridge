import time
import random
from shared_utils import GoogleCalendarAuth
from googleapiclient.errors import HttpError

# Initialize Google Calendar service
service = GoogleCalendarAuth.get_service(
    'token.json',
    'credentials.json',
    ['https://www.googleapis.com/auth/calendar']
)

CALENDAR_ID = 'brett@leathermans.net'

print('Fetching ALL events (including deleted)...')
page_token = None
all_events = []

# Fetch all events with pagination
while True:
    events_result = service.events().list(
        calendarId=CALENDAR_ID,
        showDeleted=True,
        pageToken=page_token,
        maxResults=2500
    ).execute()
    all_events.extend(events_result.get('items', []))
    page_token = events_result.get('nextPageToken')
    if not page_token:
        break

print(f'Found {len(all_events)} total events')

# Delete everything with exponential backoff
deleted_count = 0
for event in all_events:
    retries = 0
    while retries < 5:  # up to 5 retries per event
        try:
            service.events().delete(calendarId=CALENDAR_ID, eventId=event['id']).execute()
            print(f"Deleted: {event.get('summary', 'No title')}")
            deleted_count += 1
            # polite pause between calls to avoid hammering API
            time.sleep(0.1)
            break
        except HttpError as e:
            if e.resp.status == 410:
                print(f"Already deleted: {event.get('summary', 'No title')}")
                break
            elif e.resp.status in [403, 429, 500, 503]:
                # Rate limit / backend error → exponential backoff
                wait_time = (2 ** retries) + random.uniform(0, 1)
                print(f"Rate limit or server error on {event.get('summary', 'No title')}, retrying in {wait_time:.2f}s...")
                time.sleep(wait_time)
                retries += 1
            else:
                print(f"Could not delete {event.get('summary', 'No title')}: {e}")
                break
        except Exception as e:
            print(f"Unexpected error deleting {event.get('summary', 'No title')}: {e}")
            break

print(f'\nAll done! Deleted {deleted_count} events.')
