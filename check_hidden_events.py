from shared_utils import GoogleCalendarAuth
import datetime

service = GoogleCalendarAuth.get_service('token.json', 'credentials.json', ['https://www.googleapis.com/auth/calendar'])

# Check including deleted events
events_result = service.events().list(
    calendarId='brett@leathermans.net',
    showDeleted=True,
    maxResults=50
).execute()

events = events_result.get('items', [])
print(f'Total events (including deleted): {len(events)}')

for event in events[:10]:
    status = event.get('status', 'unknown')
    print(f"  Status: {status} - {event.get('summary', 'No title')} - UID: {event.get('iCalUID', 'No UID')[:30]}")
