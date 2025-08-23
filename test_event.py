import os
import datetime
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# --- Configuration ---
TOKEN_FILE = 'token.json'
SCOPES = ['https://www.googleapis.com/auth/calendar']
# !!! IMPORTANT: Use the same Calendar ID you put in your config file !!!
CALENDAR_ID = 'brett@leathermans.net'

def main():
    """Creates a single test event in the specified Google Calendar."""
    print("--- Starting Diagnostic Test ---")
    creds = None
    if not os.path.exists(TOKEN_FILE):
        print(f"ERROR: Cannot find {TOKEN_FILE}. Please run the main sync script first to authenticate.")
        return
    
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    
    try:
        service = build('calendar', 'v3', credentials=creds)
        
        # Create an event for tomorrow at 10 AM
        tomorrow = datetime.date.today() + datetime.timedelta(days=1)
        start_time = datetime.datetime(tomorrow.year, tomorrow.month, tomorrow.day, 10, 0, 0).isoformat()
        end_time = datetime.datetime(tomorrow.year, tomorrow.month, tomorrow.day, 11, 0, 0).isoformat()
        
        event = {
            'summary': 'CalendarBridge Diagnostic Test Event',
            'description': 'If you can see this, the script is working with the correct account.',
            'start': {
                'dateTime': start_time,
                'timeZone': 'America/New_York',
            },
            'end': {
                'dateTime': end_time,
                'timeZone': 'America/New_York',
            },
        }

        print(f"Attempting to create a test event in calendar: {CALENDAR_ID}")
        created_event = service.events().insert(calendarId=CALENDAR_ID, body=event).execute()
        print(f"SUCCESS! Event created. You can find it here: {created_event.get('htmlLink')}")

    except Exception as e:
        print(f"An error occurred: {e}")

if __name__ == '__main__':
    main()
