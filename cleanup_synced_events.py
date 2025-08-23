import os
import datetime
import json
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- Configuration ---
APP_DIR = os.path.dirname(os.path.realpath(__file__))
CONFIG_FILE = os.path.join(APP_DIR, 'calendar_config.json')
TOKEN_FILE = os.path.join(APP_DIR, 'token.json')
CREDENTIALS_FILE = os.path.join(APP_DIR, 'credentials.json')
SCOPES = ['https://www.googleapis.com/auth/calendar']

# How far back and forward to scan for events to delete.
DAYS_PAST = 365
DAYS_FUTURE = 365

def get_google_service():
    """Authenticates with Google and returns a Calendar service object."""
    creds = None
    if os.path.exists(TOKEN_FILE):
        creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            try:
                creds.refresh(Request())
            except Exception as e:
                print(f"Error refreshing token: {e}. Re-authenticating.")
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
        else:
            flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
            creds = flow.run_local_server(port=0)
        with open(TOKEN_FILE, 'w') as token:
            token.write(creds.to_json())
    return build('calendar', 'v3', credentials=creds)

def main():
    """Scans a Google Calendar and removes ONLY events synced from Outlook."""
    print("--- CalendarBridge Cleanup Utility (Safe Mode) ---")
    
    try:
        with open(CONFIG_FILE, 'r') as f: config = json.load(f)
        CALENDAR_ID = config.get('google_calendar_id')
        if not CALENDAR_ID:
            print(f"ERROR: 'google_calendar_id' not found in {CONFIG_FILE}")
            return
    except FileNotFoundError:
        print(f"ERROR: Configuration file '{CONFIG_FILE}' not found.")
        return

    service = get_google_service()
    
    now = datetime.datetime.utcnow()
    time_min = (now - datetime.timedelta(days=DAYS_PAST)).isoformat() + 'Z'
    time_max = (now + datetime.timedelta(days=DAYS_FUTURE)).isoformat() + 'Z'

    print(f"Scanning calendar '{CALENDAR_ID}' for Outlook-synced events to delete...")
    
    events_to_delete = []
    page_token = None
    try:
        while True:
            events_result = service.events().list(
                calendarId=CALENDAR_ID, timeMin=time_min, timeMax=time_max,
                singleEvents=True, maxResults=2500, pageToken=page_token
            ).execute()
            events = events_result.get('items', [])
            
            for event in events:
                # SAFE CHECK: Only add events to the delete list if they do NOT have a google.com iCalUID
                uid = event.get('iCalUID', '')
                if not uid.endswith('@google.com'):
                    events_to_delete.append(event)

            page_token = events_result.get('nextPageToken')
            if not page_token: break

        if not events_to_delete:
            print("Scan complete. No Outlook-synced events found to delete.")
            return

        print(f"Found {len(events_to_delete)} Outlook-synced events to delete. Your personal Google events will NOT be touched.")
        for event in events_to_delete:
             print(f"  - To be deleted: '{event.get('summary', 'No Title')}'")

        user_input = input("Do you want to proceed with deleting ONLY these events? (yes/no): ")

        if user_input.lower() != 'yes':
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
                if e.resp.status == 410: print(f"  - Already gone: '{summary}'")
                else: print(f"  - FAILED to delete '{summary}': {e}")
        
        print(f"--- Cleanup Complete. Deleted {count} events. ---")

    except HttpError as e:
        print(f"An API error occurred: {e}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

if __name__ == '__main__':
    main()
