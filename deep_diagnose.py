import os
import datetime
from shared_utils import GoogleCalendarAuth

# --- Configuration ---
APP_DIR = os.path.dirname(os.path.realpath(__file__))
TOKEN_FILE = os.path.join(APP_DIR, 'token.json')
CREDENTIALS_FILE = os.path.join(APP_DIR, 'credentials.json')
SCOPES = ['https://www.googleapis.com/auth/calendar']
CALENDAR_ID = 'brett@leathermans.net'

def main():
    """Performs a deep diagnosis of the Google Calendar."""
    print("--- Starting CalendarBridge Deep Diagnostics ---")
    
    try:
        service = GoogleCalendarAuth.get_service(TOKEN_FILE, CREDENTIALS_FILE, SCOPES)
        
        # Define the exact sync window from your logs
        time_min = "2025-08-24T00:00:00-04:00"
        time_max = "2025-08-31T23:59:59-04:00"
        
        print(f"\n--- Checking for events in the sync window ({time_min} to {time_max}) ---")
        print(f"Including deleted/cancelled events...")

        events_result = service.events().list(
            calendarId=CALENDAR_ID,
            timeMin=time_min,
            timeMax=time_max,
            maxResults=50,
            singleEvents=True,
            showDeleted=True,
            orderBy='startTime'
        ).execute()
        events = events_result.get('items', [])

        if not events:
            print("\n>>> RESULT: No events (including hidden ones) were found in the sync window.")
            print("This confirms the events are not being created successfully, despite the API's response.")
        else:
            print(f"\n>>> RESULT: Found {len(events)} events in the sync window!")
            for event in events:
                summary = event.get('summary', 'No Title')
                start = event.get('start', {}).get('dateTime', event.get('start', {}).get('date'))
                status = event.get('status', 'No Status')
                uid = event.get('iCalUID', 'No UID')
                creator = event.get('creator', {}).get('email', 'Unknown')
                
                print("\n----------------------------------------")
                print(f"  Event: {summary}")
                print(f"  Start: {start}")
                print(f"  Status: {status}")
                print(f"  UID: {uid}")
                print(f"  Creator: {creator}")
                print("----------------------------------------")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    main()
