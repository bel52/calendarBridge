import os
import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- Configuration ---
# This is the ID of the calendar you want to clean.
CALENDAR_ID = 'brett@leathermans.net'

# How far back and forward to scan for events to delete.
DAYS_PAST = 365
DAYS_FUTURE = 365

# --- Do not edit below this line ---
TOKEN_FILE = 'token.json'
SCOPES = ['https://www.googleapis.com/auth/calendar']

def main():
    """Scans a Google Calendar and removes events imported from non-Google sources."""
    print("--- CalendarBridge Cleanup Utility (v2) ---")
    
    creds = None
    if not os.path.exists(TOKEN_FILE):
        print(f"ERROR: Authentication file '{TOKEN_FILE}' not found.")
        print("Please run the main sync.sh script at least once to create it.")
        return
        
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    service = build('calendar', 'v3', credentials=creds)
    
    now = datetime.datetime.utcnow()
    time_min = (now - datetime.timedelta(days=DAYS_PAST)).isoformat() + 'Z'
    time_max = (now + datetime.timedelta(days=DAYS_FUTURE)).isoformat() + 'Z'

    print(f"Scanning calendar '{CALENDAR_ID}' for imported events...")
    
    try:
        events_result = service.events().list(
            calendarId=CALENDAR_ID,
            timeMin=time_min,
            timeMax=time_max,
            singleEvents=True,
            orderBy='startTime',
            maxResults=2500 # Get as many results as possible
        ).execute()
        events = events_result.get('items', [])

        if not events:
            print("No events found in the specified date range. All clean!")
            return

        events_to_delete = []
        for event in events:
            # Get the iCalUID, which is the unique identifier from the source system.
            uid = event.get('iCalUID', '')
            
            # Events created natively in Google Calendar have a UID ending in @google.com
            # Events imported from Outlook/Exchange have a different format.
            # This is the most reliable way to distinguish them.
            if not uid.endswith('@google.com'):
                events_to_delete.append(event)

        if not events_to_delete:
            print("Scan complete. No synced Outlook events found to delete.")
            return

        print(f"Found {len(events_to_delete)} potentially synced events.")
        user_input = input("Do you want to proceed with deleting them? (yes/no): ").lower().strip()

        # FIX: Accept 'y' or 'yes'
        if user_input not in ['y', 'yes']:
            print("Aborting. No events were deleted.")
            return

        print("--- Starting Deletion ---")
        count = 0
        for event in events_to_delete:
            summary = event.get('summary', 'No Title')
            event_id = event['id']
            try:
                service.events().delete(calendarId=CALENDAR_ID, eventId=event_id).execute()
                print(f"  - Deleted: '{summary}'")
                count += 1
            except HttpError as e:
                print(f"  - FAILED to delete '{summary}': {e}")
        
        print(f"--- Cleanup Complete. Deleted {count} events. ---")

    except HttpError as e:
        print(f"An API error occurred: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == '__main__':
    main()
