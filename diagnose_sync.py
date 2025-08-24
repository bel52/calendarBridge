import os
from shared_utils import GoogleCalendarAuth

# --- Configuration ---
APP_DIR = os.path.dirname(os.path.realpath(__file__))
TOKEN_FILE = os.path.join(APP_DIR, 'token.json')
CREDENTIALS_FILE = os.path.join(APP_DIR, 'credentials.json')
SCOPES = ['https://www.googleapis.com/auth/calendar']
CALENDAR_ID = 'brett@leathermans.net'

def main():
    """Diagnoses the state of the Google Calendar."""
    print("--- Starting CalendarBridge Diagnostics ---")
    
    try:
        service = GoogleCalendarAuth.get_service(TOKEN_FILE, CREDENTIALS_FILE, SCOPES)
        
        # 1. List all available calendars
        print("\n--- 1. Available Calendars ---")
        calendars_result = service.calendarList().list().execute()
        calendars = calendars_result.get('items', [])

        if not calendars:
            print("No calendars found.")
        else:
            print("Your account has access to the following calendars:")
            for calendar in calendars:
                summary = calendar['summary']
                cal_id = calendar['id']
                access = calendar['accessRole']
                print(f"  - Summary: {summary}")
                print(f"    ID: {cal_id}")
                print(f"    Access Role: {access}")
                if cal_id == CALENDAR_ID:
                    print("    (This is your target calendar)")

        # 2. Check for recent events on the target calendar
        print(f"\n--- 2. Checking for recent events on '{CALENDAR_ID}' ---")
        events_result = service.events().list(
            calendarId=CALENDAR_ID,
            maxResults=10,
            singleEvents=True,
            orderBy='startTime'
        ).execute()
        events = events_result.get('items', [])

        if not events:
            print("No upcoming events found on this calendar.")
        else:
            print(f"Found {len(events)} upcoming events:")
            for event in events:
                summary = event.get('summary', 'No Title')
                start = event.get('start', {}).get('dateTime', event.get('start', {}).get('date'))
                status = event.get('status', 'No Status')
                uid = event.get('iCalUID', 'No UID')
                print(f"  - Summary: {summary}")
                print(f"    Start: {start}")
                print(f"    Status: {status}")
                print(f"    UID: {uid}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    main()
