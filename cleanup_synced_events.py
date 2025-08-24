import os
import datetime
import json
from googleapiclient.errors import HttpError
from shared_utils import GoogleCalendarAuth, ConfigManager

# --- Configuration ---
APP_DIR = os.path.dirname(os.path.realpath(__file__))
CONFIG_FILE = os.path.join(APP_DIR, 'calendar_config.json')
TOKEN_FILE = os.path.join(APP_DIR, 'token.json')
CREDENTIALS_FILE = os.path.join(APP_DIR, 'credentials.json')
SCOPES = ['https://www.googleapis.com/auth/calendar']

# How far back and forward to scan for events to delete.
DAYS_PAST = 365
DAYS_FUTURE = 365

def main():
    """Scans a Google Calendar and removes ONLY events synced from Outlook."""
    print("--- CalendarBridge Cleanup Utility (Enhanced) ---")
    
    try:
        config = ConfigManager.load_config(CONFIG_FILE)
        CALENDAR_ID = config.get('google_calendar_id')
        if not CALENDAR_ID:
            print(f"ERROR: 'google_calendar_id' not found in {CONFIG_FILE}")
            return
    except FileNotFoundError:
        print(f"ERROR: Configuration file '{CONFIG_FILE}' not found.")
        return
    except ValueError as e:
        print(f"ERROR: Configuration validation failed: {e}")
        return

    service = GoogleCalendarAuth.get_service(TOKEN_FILE, CREDENTIALS_FILE, SCOPES)
    
    now = datetime.datetime.utcnow()
    time_min = (now - datetime.timedelta(days=DAYS_PAST)).isoformat() + 'Z'
    time_max = (now + datetime.timedelta(days=DAYS_FUTURE)).isoformat() + 'Z'

    print(f"Scanning calendar '{CALENDAR_ID}' for Outlook-synced events...")
    print(f"Scan window: {DAYS_PAST} days past to {DAYS_FUTURE} days future")
    
    events_to_delete = []
    page_token = None
    total_events_scanned = 0
    
    try:
        while True:
            events_result = service.events().list(
                calendarId=CALENDAR_ID, timeMin=time_min, timeMax=time_max,
                singleEvents=True, maxResults=2500, pageToken=page_token
            ).execute()
            events = events_result.get('items', [])
            total_events_scanned += len(events)
            
            for event in events:
                uid = event.get('iCalUID', '')
                # Check for CalendarBridge UIDs or non-Google UIDs
                if uid.endswith('@cbridge.local') or (uid and not uid.endswith('@google.com')):
                    events_to_delete.append(event)

            page_token = events_result.get('nextPageToken')
            if not page_token: 
                break

        print(f"Scanned {total_events_scanned} total events")
        
        if not events_to_delete:
            print("✓ No Outlook-synced events found to delete. Calendar is clean!")
            return

        print(f"\n⚠️  Found {len(events_to_delete)} Outlook-synced events to delete.")
        print("Your personal Google Calendar events will NOT be touched.\n")
        
        # Show first 5 events as preview
        print("Preview of events to be deleted:")
        for i, event in enumerate(events_to_delete[:5]):
            start = event.get('start', {})
            date_str = start.get('date', start.get('dateTime', 'Unknown date'))
            print(f"  • {event.get('summary', 'No Title')} ({date_str})")
        if len(events_to_delete) > 5:
            print(f"  ... and {len(events_to_delete) - 5} more events")

        print("\n" + "="*50)
        user_input = input("Do you want to proceed with deleting ONLY these events? (yes/no): ")

        if user_input.lower() not in ['yes', 'y']:
            print("❌ Aborted. No events were deleted.")
            return

        print("\n--- Starting Deletion ---")
        count = 0
        failed = 0
        
        for event in events_to_delete:
            summary = event.get('summary', 'No Title')
            event_id = event['id']
            try:
                service.events().delete(calendarId=CALENDAR_ID, eventId=event_id).execute()
                print(f"  ✓ Deleted: '{summary}'")
                count += 1
            except HttpError as e:
                if e.resp.status == 410: 
                    print(f"  ⚠ Already gone: '{summary}'")
                else: 
                    print(f"  ✗ FAILED to delete '{summary}': {e}")
                    failed += 1
        
        print("\n" + "="*50)
        print(f"✅ Cleanup Complete!")
        print(f"   • Deleted: {count} events")
        if failed > 0:
            print(f"   • Failed: {failed} events")
        print("="*50)

    except HttpError as e:
        print(f"❌ An API error occurred: {e}")
    except Exception as e:
        print(f"❌ An unexpected error occurred: {e}")

if __name__ == '__main__':
    main()
