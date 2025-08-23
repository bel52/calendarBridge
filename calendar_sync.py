import os
import json
import logging
from logging.handlers import RotatingFileHandler
import datetime
import subprocess
import pytz
import time
import random
import hashlib
from icalendar import Calendar
import recurring_ical_events
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# --- Configuration ---
APP_DIR = os.path.dirname(os.path.realpath(__file__))
OUTBOX_DIR = os.path.join(APP_DIR, 'outbox')
EXPORT_FILE = os.path.join(OUTBOX_DIR, 'outlook_full_export.ics')
CONFIG_FILE = os.path.join(APP_DIR, 'calendar_config.json')
TOKEN_FILE = os.path.join(APP_DIR, 'token.json')
CREDENTIALS_FILE = os.path.join(APP_DIR, 'credentials.json')
LOG_FILE = os.path.join(APP_DIR, 'logs', 'calendar_sync.log')
SCOPES = ['https://www.googleapis.com/auth/calendar']
TIMEZONE = "America/New_York"

# --- Logger Setup ---
def setup_logger():
    """Initializes and configures the rotating file logger."""
    os.makedirs(os.path.join(APP_DIR, 'logs'), exist_ok=True)
    logger = logging.getLogger("CalendarBridge")
    if logger.hasHandlers(): logger.handlers.clear()
    logger.setLevel(logging.INFO)
    handler = RotatingFileHandler(LOG_FILE, maxBytes=5*1024*1024, backupCount=5)
    formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
    handler.setFormatter(formatter)
    logger.addHandler(handler)
    logger.addHandler(logging.StreamHandler())
    return logger

logger = setup_logger()

# --- Main Application Logic ---
class CalendarSync:
    def __init__(self):
        self.config = self.load_config()
        self.google_service = self.get_google_service()
        self.tz = pytz.timezone(TIMEZONE)

    def load_config(self):
        """Loads the configuration from the JSON file."""
        with open(CONFIG_FILE, 'r') as f: return json.load(f)

    def get_google_service(self):
        """Authenticates with Google and returns a Calendar service object."""
        creds = None
        if os.path.exists(TOKEN_FILE):
            creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception as e:
                    logger.error(f"Error refreshing token: {e}. Re-authenticating.")
                    flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                    creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_FILE, SCOPES)
                creds = flow.run_local_server(port=0)
            with open(TOKEN_FILE, 'w') as token:
                token.write(creds.to_json())
        return build('calendar', 'v3', credentials=creds)

    def execute_with_backoff(self, api_call):
        """Executes a Google API call with exponential backoff for retries."""
        max_retries = 5
        for attempt in range(max_retries):
            try:
                return api_call.execute()
            except HttpError as e:
                if e.resp.status == 409:
                    logger.warning(f"Ignoring 409 error for what is likely a duplicate event.")
                    return None
                if e.resp.status in [403, 500, 503]:
                    if attempt < max_retries - 1:
                        wait_time = (2 ** attempt) + random.uniform(0, 1)
                        logger.warning(f"Rate limit or server error. Retrying in {wait_time:.2f} seconds...")
                        time.sleep(wait_time)
                    else:
                        logger.error("Max retries exceeded.")
                        raise
                elif e.resp.status in [404, 410]:
                    logger.warning("Resource not found or already gone. Ignoring.")
                    return None
                else:
                    raise
        return None

    def run_applescript_export(self):
        """Triggers the AppleScript to export the Outlook calendar to an ICS file."""
        logger.info("Triggering AppleScript for Outlook export...")
        script_path = os.path.join(APP_DIR, 'exportEvents.scpt')
        outlook_calendar_name = self.config.get("outlook_calendar_name")
        outlook_calendar_index = self.config.get("outlook_calendar_index", 1)

        if not outlook_calendar_name:
            logger.error("outlook_calendar_name not set in calendar_config.json")
            return False
        try:
            os.makedirs(OUTBOX_DIR, exist_ok=True)
            result = subprocess.run(
                ['osascript', script_path, outlook_calendar_name, str(outlook_calendar_index)],
                check=True, capture_output=True, text=True
            )
            logger.info("AppleScript successfully triggered the export.")
            return True
        except subprocess.CalledProcessError as e:
            logger.error(f"AppleScript execution failed:\nSTDOUT: {e.stdout}\nSTDERR: {e.stderr}")
            return False

    def get_existing_google_events(self, calendar_id, time_min, time_max):
        """Fetches existing events from Google Calendar within the sync window."""
        logger.info("Fetching existing Google Calendar events...")
        all_events = {}
        page_token = None
        while True:
            events_result = self.google_service.events().list(
                calendarId=calendar_id, timeMin=time_min.isoformat(), timeMax=time_max.isoformat(),
                singleEvents=True, showDeleted=False, maxResults=2500, pageToken=page_token
            ).execute()
            for event in events_result.get('items', []):
                if 'iCalUID' in event:
                    all_events[event['iCalUID']] = event
            page_token = events_result.get('nextPageToken')
            if not page_token:
                break
        logger.info(f"Found {len(all_events)} existing events in Google Calendar.")
        return all_events

    def build_event_body(self, component, uid):
        """Builds the Google Calendar event body from an ical component."""
        summary = str(component.get('summary', 'No Title'))
        dtstart = component.get('dtstart').dt
        dtend = component.get('dtend').dt

        is_all_day = False
        is_standard_all_day = not isinstance(dtstart, datetime.datetime)

        if is_standard_all_day:
            is_all_day = True
        elif isinstance(dtstart, datetime.datetime):
            starts_at_midnight = dtstart.time() == datetime.time(0, 0)
            duration = dtend - dtstart
            is_full_day_duration = duration.total_seconds() > 0 and duration.total_seconds() % (24 * 3600) == 0
            if starts_at_midnight and is_full_day_duration:
                is_all_day = True

        if is_all_day:
            start_date = dtstart if is_standard_all_day else dtstart.date()
            end_date = dtend if is_standard_all_day else dtend.date()
            start = {'date': start_date.strftime('%Y-%m-%d')}
            end = {'date': end_date.strftime('%Y-%m-%d')}
        else:
            start_dt = dtstart if dtstart.tzinfo else self.tz.localize(dtstart)
            end_dt = dtend if dtend.tzinfo else self.tz.localize(end_dt)
            start = {'dateTime': start_dt.isoformat(), 'timeZone': TIMEZONE}
            end = {'dateTime': end_dt.isoformat(), 'timeZone': TIMEZONE}

        event_body = {'summary': summary, 'start': start, 'end': end, 'iCalUID': uid}
        
        if component.get('location'):
            event_body['location'] = str(component.get('location'))
        if component.get('description'):
            event_body['description'] = str(component.get('description'))
        
        transp = component.get('transp')
        if transp and str(transp) == 'TRANSPARENT':
            event_body['transparency'] = 'transparent'
        
        return event_body

    def sync_events(self):
        """The core logic to parse the ICS file and sync events to Google Calendar."""
        target_calendar_id = self.config.get('google_calendar_id')
        days_past = self.config.get('sync_days_past', 90)
        days_future = self.config.get('sync_days_future', 120)
        api_delay = self.config.get('api_delay_seconds', 0.1)
        
        now = self.tz.localize(datetime.datetime.now())
        time_min = now - datetime.timedelta(days=days_past)
        time_max = now + datetime.timedelta(days=days_future)
        logger.info(f"Sync window is from {time_min.strftime('%Y-%m-%d')} to {time_max.strftime('%Y-%m-%d')}")

        if not os.path.exists(EXPORT_FILE):
            logger.error(f"Export file not found at {EXPORT_FILE}. Aborting sync."); return

        try:
            with open(EXPORT_FILE, 'r', encoding='utf-8') as f: ical_data = f.read()
        except UnicodeDecodeError:
            with open(EXPORT_FILE, 'r', encoding='latin-1') as f: ical_data = f.read()
        
        if not ical_data.strip():
            logger.warning("ICS file is empty. Cleaning up and finishing cycle."); os.remove(EXPORT_FILE); return

        all_expanded_events = []
        calendar_sections = ical_data.split('BEGIN:VCALENDAR')
        
        for section in calendar_sections:
            if not section.strip(): continue
            calendar_str = 'BEGIN:VCALENDAR' + section
            try:
                cal = Calendar.from_ical(calendar_str)
                expanded = recurring_ical_events.of(cal).between(time_min, time_max)
                all_expanded_events.extend(expanded)
            except Exception as e:
                logger.warning(f"Could not parse a section of the ICS file: {e}")
            
        existing_google_events = self.get_existing_google_events(target_calendar_id, time_min, time_max)
        outlook_uids_in_range = set()
        processed_uids_this_run = set()

        logger.info(f"Processing {len(all_expanded_events)} expanded event instances from the ICS file...")
        for component in all_expanded_events:
            try:
                summary = str(component.get('summary', 'No Title'))
                dtstart = component.get('dtstart').dt

                # --- New, Definitive UID Generation ---
                is_all_day = not isinstance(dtstart, datetime.datetime)
                if is_all_day:
                    uid_timestamp = dtstart.strftime('%Y%m%d')
                else:
                    start_dt_for_uid = dtstart if dtstart.tzinfo else self.tz.localize(dtstart)
                    uid_timestamp = start_dt_for_uid.astimezone(pytz.utc).strftime('%Y%m%dT%H%M%SZ')
                
                # Create a unique, stable fingerprint for the event instance.
                event_signature = f"{summary}{uid_timestamp}".encode('utf-8')
                hashed_signature = hashlib.sha1(event_signature).hexdigest()
                uid = f"{hashed_signature}@cbridge.local"
                # --- End of New UID Generation ---

                outlook_uids_in_range.add(uid)

                if uid in processed_uids_this_run:
                    continue
                
                new_event_body = self.build_event_body(component, uid)

                if uid not in existing_google_events:
                    logger.info(f"CREATING event: '{summary}'")
                    api_request = self.google_service.events().insert(calendarId=target_calendar_id, body=new_event_body)
                    self.execute_with_backoff(api_request)
                    time.sleep(api_delay)
                else:
                    existing_event = existing_google_events[uid]
                    is_changed = False
                    # Only need to compare summary now as other fields are part of the UID
                    if new_event_body.get('summary', '') != existing_event.get('summary', ''):
                        is_changed = True
                    
                    if is_changed:
                        logger.info(f"UPDATING event: '{summary}'")
                        event_id = existing_event['id']
                        api_request = self.google_service.events().update(calendarId=target_calendar_id, eventId=event_id, body=new_event_body)
                        self.execute_with_backoff(api_request)
                        time.sleep(api_delay)
                
                processed_uids_this_run.add(uid)

            except Exception as e:
                logger.error(f"An unexpected error occurred while processing event '{component.get('summary', 'Unknown Event')}': {e}", exc_info=True)
        
        google_uids_in_range = set(existing_google_events.keys())
        orphan_uids = google_uids_in_range - outlook_uids_in_range

        if orphan_uids:
            logger.info(f"--- Deleting {len(orphan_uids)} Orphan Events ---")
            for uid in orphan_uids:
                event_to_delete = existing_google_events[uid]
                summary = event_to_delete.get('summary', 'No Title')
                event_id = event_to_delete['id']
                logger.info(f"DELETING orphan event: {summary} (UID: {uid})")
                api_request = self.google_service.events().delete(calendarId=target_calendar_id, eventId=event_id)
                self.execute_with_backoff(api_request)
                time.sleep(api_delay)
        
        if os.path.exists(EXPORT_FILE):
            os.remove(EXPORT_FILE)

    def run(self):
        """Main execution flow for a single sync cycle."""
        logger.info("--- Starting CalendarBridge Sync Cycle (v3.7 - Generated UIDs) ---")
        if self.run_applescript_export():
            wait_start_time = time.time()
            max_wait_seconds = 60
            while not os.path.exists(EXPORT_FILE):
                time.sleep(1)
                if time.time() - wait_start_time > max_wait_seconds:
                    logger.error(f"Timeout: Export file was not created within {max_wait_seconds} seconds.")
                    return

            logger.info("Export file found. Proceeding with sync.")
            self.sync_events()
        logger.info("--- Sync Cycle Finished ---")

if __name__ == '__main__':
    sync_app = CalendarSync()
    sync_app.run()
