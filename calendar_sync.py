#!/usr/bin/env python3
import os, json, logging, datetime, subprocess, time, pytz, random, hashlib
from logging.handlers import RotatingFileHandler
from icalendar import Calendar
import recurring_ical_events
from google.oauth2.credentials import Credentials
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

class CalendarSync:
    def __init__(self):
        self.app_dir = os.path.dirname(os.path.realpath(__file__))
        self.config_path = os.path.join(self.app_dir, 'calendar_config.json')
        self.export_file = os.path.join(self.app_dir, 'outbox', 'outlook_full_export.ics')
        self.token_file = os.path.join(self.app_dir, 'token.json')
        self.credentials_file = os.path.join(self.app_dir, 'credentials.json')
        self.log_file = os.path.join(self.app_dir, 'logs', 'calendar_sync.log')
        self.config = self.load_config()
        # Use timezone from config if provided, otherwise default to America/New_York
        self.tz = pytz.timezone(self.config.get('timezone', 'America/New_York'))
        self.logger = self.setup_logger()
        self.google_service = self.get_google_service()

    def setup_logger(self):
        """Set up a rotating file logger and a console logger."""
        logger = logging.getLogger('CalendarBridge')
        if logger.hasHandlers():
            logger.handlers.clear()
        logger.setLevel(logging.INFO)
        os.makedirs(os.path.join(self.app_dir, 'logs'), exist_ok=True)
        handler = RotatingFileHandler(self.log_file, maxBytes=5 * 1024 * 1024, backupCount=5)
        formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        handler.setFormatter(formatter)
        logger.addHandler(handler)
        logger.addHandler(logging.StreamHandler())
        return logger

    def load_config(self):
        """Load JSON configuration from calendar_config.json."""
        try:
            with open(self.config_path, 'r') as f:
                return json.load(f)
        except Exception as e:
            raise RuntimeError(f'Error loading config file {self.config_path}: {e}')

    def get_google_service(self):
        """Authenticate with Google Calendar and return a service object."""
        creds = None
        scopes = ['https://www.googleapis.com/auth/calendar']
        if os.path.exists(self.token_file):
            creds = Credentials.from_authorized_user_file(self.token_file, scopes)
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                try:
                    creds.refresh(Request())
                except Exception as e:
                    self.logger.error(f'Error refreshing token: {e}. Re-authenticating.')
                    flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, scopes)
                    creds = flow.run_local_server(port=0)
            else:
                flow = InstalledAppFlow.from_client_secrets_file(self.credentials_file, scopes)
                creds = flow.run_local_server(port=0)
            with open(self.token_file, 'w') as token:
                token.write(creds.to_json())
        return build('calendar', 'v3', credentials=creds)

    def run_applescript_export(self):
        """Trigger the AppleScript to export Outlook events to outbox/outlook_full_export.ics."""
        outlook_calendar_name = self.config.get('outlook_calendar_name')
        outlook_calendar_index = self.config.get('outlook_calendar_index', 1)
        if not outlook_calendar_name:
            self.logger.error("outlook_calendar_name not set in config")
            return False
        script_path = os.path.join(self.app_dir, 'exportEvents.scpt')
        os.makedirs(os.path.join(self.app_dir, 'outbox'), exist_ok=True)
        self.logger.info("Triggering AppleScript for Outlook export...")
        try:
            subprocess.run(
                ['osascript', script_path, outlook_calendar_name, str(outlook_calendar_index)],
                check=True, capture_output=True, text=True
            )
            self.logger.info("AppleScript successfully triggered the export.")
            return True
        except subprocess.CalledProcessError as e:
            self.logger.error(f"AppleScript execution failed:\nSTDOUT: {e.stdout}\nSTDERR: {e.stderr}")
            return False

    def clean_ics_data(self, ical_data):
        """
        Remove non-standard X- headers that can confuse the icalendar parser.
        This is equivalent to the previous clean_ics_files.py functionality.
        """
        cleaned_lines = []
        for line in ical_data.splitlines():
            # Skip lines starting with certain X- prefixes known to be problematic
            if line.startswith('X-ENTOURAGE_UUID') or line.startswith('X-CALENDARSERVER-') or line.startswith('X-MS-'):
                continue
            cleaned_lines.append(line)
        return "\n".join(cleaned_lines)

    def parse_ics_file(self, path, time_min, time_max):
        """Load the .ics file, clean it, and expand recurring events within the sync window."""
        try:
            with open(path, 'r', encoding='utf-8') as f:
                ical_data = f.read()
        except UnicodeDecodeError:
            with open(path, 'r', encoding='latin-1') as f:
                ical_data = f.read()
        if not ical_data.strip():
            return []
        ical_data = self.clean_ics_data(ical_data)
        expanded_events = []
        sections = ical_data.split('BEGIN:VCALENDAR')
        for section in sections:
            if not section.strip():
                continue
            calendar_str = 'BEGIN:VCALENDAR' + section
            try:
                cal = Calendar.from_ical(calendar_str)
                events = recurring_ical_events.of(cal).between(time_min, time_max)
                expanded_events.extend(events)
            except Exception as e:
                self.logger.warning(f"Could not parse a section of the ics file: {e}")
        return expanded_events

    def get_existing_google_events(self, calendar_id, time_min, time_max):
        """
        Fetch existing events in Google Calendar within the sync window.
        Returns a dict keyed by iCalUID for quick lookup.
        """
        existing = {}
        page_token = None
        while True:
            events_result = self.google_service.events().list(
                calendarId=calendar_id,
                timeMin=time_min.isoformat(),
                timeMax=time_max.isoformat(),
                singleEvents=True,
                showDeleted=False,
                maxResults=2500,
                pageToken=page_token
            ).execute()
            for event in events_result.get('items', []):
                if 'iCalUID' in event:
                    existing[event['iCalUID']] = event
            page_token = events_result.get('nextPageToken')
            if not page_token:
                break
        return existing

    def compute_uid(self, component):
        """
        Compute a stable UID. If the original Outlook UID exists, use it as the base; otherwise use the summary.
        Combine it with a timestamp derived from the event start (in UTC) or date-only for all-day events, then hash.
        """
        summary = str(component.get('summary', 'No Title'))
        dtstart = component.get('dtstart').dt
        # Determine if all-day
        is_all_day = not isinstance(dtstart, datetime.datetime)
        if is_all_day:
            uid_timestamp = dtstart.strftime('%Y%m%d')
        else:
            start_dt = dtstart if dtstart.tzinfo else self.tz.localize(dtstart)
            uid_timestamp = start_dt.astimezone(pytz.utc).strftime('%Y%m%dT%H%M%SZ')
        source_uid = ''
        try:
            source_uid = str(component.get('uid'))
        except Exception:
            source_uid = ''
        base = source_uid if source_uid else summary
        signature = f"{base}{uid_timestamp}".encode('utf-8')
        hashed = hashlib.sha1(signature).hexdigest()
        return f"{hashed}@cbridge.local"

    def build_event_body(self, component, uid):
        """
        Build the request body for Google Calendar. Handles all-day events, timed events,
        and optional fields like location and description. Transparently sets free/busy.
        """
        summary = str(component.get('summary', 'No Title'))
        dtstart = component.get('dtstart').dt
        dtend = component.get('dtend').dt
        # Decide if it's an all-day event
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
            end_dt = dtend if dtend.tzinfo else self.tz.localize(dtend)
            start = {'dateTime': start_dt.isoformat(), 'timeZone': self.tz.zone}
            end = {'dateTime': end_dt.isoformat(), 'timeZone': self.tz.zone}
        event_body = {
            'summary': summary,
            'start': start,
            'end': end,
            'iCalUID': uid
        }
        if component.get('location'):
            event_body['location'] = str(component.get('location'))
        if component.get('description'):
            event_body['description'] = str(component.get('description'))
        transp = component.get('transp')
        if transp and str(transp) == 'TRANSPARENT':
            event_body['transparency'] = 'transparent'
        return event_body

    def sync_events(self):
        """Core sync logic: create/update/delete events to mirror Outlook."""
        calendar_id = self.config.get('google_calendar_id')
        days_past = self.config.get('sync_days_past', 90)
        days_future = self.config.get('sync_days_future', 120)
        api_delay = self.config.get('api_delay_seconds', 0.1)
        now = self.tz.localize(datetime.datetime.now())
        time_min = now - datetime.timedelta(days=days_past)
        time_max = now + datetime.timedelta(days=days_future)
        if not os.path.exists(self.export_file):
            self.logger.error(f"Export file not found at {self.export_file}. Aborting sync.")
            return
        # Parse and expand events from Outlook
        expanded_events = self.parse_ics_file(self.export_file, time_min, time_max)
        existing_google_events = self.get_existing_google_events(calendar_id, time_min, time_max)
        outlook_uids = set()
        processed_uids = set()
        for component in expanded_events:
            try:
                uid = self.compute_uid(component)
                outlook_uids.add(uid)
                # Skip duplicate instances within the same run
                if uid in processed_uids:
                    continue
                event_body = self.build_event_body(component, uid)
                if uid not in existing_google_events:
                    # Create new event
                    self.logger.info(f"Creating: {event_body['summary']}")
                    request = self.google_service.events().insert(calendarId=calendar_id, body=event_body)
                    self.execute_with_backoff(request)
                    time.sleep(api_delay)
                else:
                    # Compare for updates (only compare fields that can change independently of start/end)
                    existing_event = existing_google_events[uid]
                    is_changed = False
                    for key in ('summary', 'location', 'description'):
                        if event_body.get(key, '') != existing_event.get(key, ''):
                            is_changed = True
                            break
                    if is_changed:
                        self.logger.info(f"Updating: {event_body['summary']}")
                        event_id = existing_event['id']
                        request = self.google_service.events().update(calendarId=calendar_id, eventId=event_id, body=event_body)
                        self.execute_with_backoff(request)
                        time.sleep(api_delay)
                processed_uids.add(uid)
            except Exception as e:
                # Log and continue with other events
                self.logger.error(f"Unexpected error processing event '{component.get('summary','Unknown Event')}': {e}", exc_info=True)
        # Delete orphan events: events in Google but not in Outlook
        google_uids = set(existing_google_events.keys())
        orphan_uids = google_uids - outlook_uids
        for uid in orphan_uids:
            try:
                event_to_delete = existing_google_events[uid]
                summary = event_to_delete.get('summary', 'No Title')
                event_id = event_to_delete['id']
                self.logger.info(f"Deleting orphan: {summary} (UID: {uid})")
                request = self.google_service.events().delete(calendarId=calendar_id, eventId=event_id)
                self.execute_with_backoff(request)
                time.sleep(api_delay)
            except Exception as e:
                self.logger.error(f"Failed to delete orphan event {uid}: {e}", exc_info=True)
        # Clean up export file
        if os.path.exists(self.export_file):
            os.remove(self.export_file)

    def execute_with_backoff(self, api_request):
        """
        Execute a Google Calendar API request with exponential backoff to handle
        transient errors and rate limits.
        """
        max_retries = 5
        for attempt in range(max_retries):
            try:
                return api_request.execute()
            except HttpError as e:
                status = e.resp.status
                if status == 409:
                    # Conflict: likely duplicate. Safe to ignore.
                    self.logger.warning("Conflict (likely duplicate). Ignoring.")
                    return None
                if status in [403, 500, 503]:
                    # Respect retry-after recommendations
                    if attempt < max_retries - 1:
                        delay = (2 ** attempt) + random.uniform(0, 1)
                        self.logger.warning(f"API rate limit or server error. Retrying in {delay:.2f} seconds...")
                        time.sleep(delay)
                    else:
                        self.logger.error("Max retries exceeded.")
                        raise
                elif status in [404, 410]:
                    # Resource not found or already deleted
                    self.logger.warning("Resource not found or already deleted.")
                    return None
                else:
                    raise

    def run(self):
        """Top-level run method: export, wait for file, and sync."""
        self.logger.info("--- Starting CalendarBridge Sync Cycle ---")
        if self.run_applescript_export():
            # Wait up to 60 seconds for the export file to appear
            start_time = time.time()
            max_wait = 60
            while not os.path.exists(self.export_file):
                time.sleep(1)
                if time.time() - start_time > max_wait:
                    self.logger.error(f"Timeout: export file not created within {max_wait} seconds.")
                    return
            self.logger.info("Export file found. Proceeding with sync.")
            self.sync_events()
        self.logger.info("--- Sync Cycle Finished ---")

if __name__ == '__main__':
    sync_app = CalendarSync()
    sync_app.run()
