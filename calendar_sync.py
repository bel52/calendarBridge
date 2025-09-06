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
        self.tz = pytz.timezone(self.config.get('timezone', 'America/New_York'))
        self.enable_batch = bool(self.config.get('enable_batch_operations', False))
        self.batch_size = int(self.config.get('batch_size', 50))
        self.api_delay = float(self.config.get('api_delay_seconds', 0.1))

        self.logger = self.setup_logger()
        self.google_service = self.get_google_service()

    # ---------- infra ----------

    def setup_logger(self):
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
        with open(self.config_path, 'r') as f:
            return json.load(f)

    def get_google_service(self):
        scopes = ['https://www.googleapis.com/auth/calendar']
        creds = None
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

    # ---------- Outlook export ----------

    def run_applescript_export(self):
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

    # ---------- ICS handling ----------

    def clean_ics_data(self, ical_data: str) -> str:
        cleaned = []
        for line in ical_data.splitlines():
            if line.startswith('X-ENTOURAGE_UUID') or line.startswith('X-CALENDARSERVER-') or line.startswith('X-MS-'):
                continue
            cleaned.append(line)
        return "\n".join(cleaned)

    def parse_ics_file(self, path, time_min, time_max):
        try:
            try:
                with open(path, 'r', encoding='utf-8') as f:
                    ical_data = f.read()
            except UnicodeDecodeError:
                with open(path, 'r', encoding='latin-1') as f:
                    ical_data = f.read()
        except FileNotFoundError:
            return []
        if not ical_data.strip():
            return []
        ical_data = self.clean_ics_data(ical_data)

        expanded_events = []
        sections = ical_data.split('BEGIN:VCALENDAR')
        for section in sections:
            s = section.strip()
            if not s:
                continue
            try:
                cal = Calendar.from_ical('BEGIN:VCALENDAR' + section)
                events = recurring_ical_events.of(cal).between(time_min, time_max)
                expanded_events.extend(events)
            except Exception as e:
                self.logger.warning(f"Could not parse one VCALENDAR: {e}")
        return expanded_events

    # ---------- Google helpers ----------

    def get_existing_google_events(self, calendar_id, time_min, time_max):
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
        summary = str(component.get('summary', 'No Title'))
        dtstart = component.get('dtstart').dt
        is_all_day = not isinstance(dtstart, datetime.datetime)
        if is_all_day:
            uid_ts = dtstart.strftime('%Y%m%d')
        else:
            start_dt = dtstart if dtstart.tzinfo else self.tz.localize(dtstart)
            uid_ts = start_dt.astimezone(pytz.utc).strftime('%Y%m%dT%H%M%SZ')
        try:
            source_uid = str(component.get('uid')) or ''
        except Exception:
            source_uid = ''
        base = source_uid if source_uid else summary
        hashed = hashlib.sha1(f"{base}{uid_ts}".encode('utf-8')).hexdigest()
        return f"{hashed}@cbridge.local"

    def build_event_body(self, component, uid):
        summary = str(component.get('summary', 'No Title'))
        dtstart = component.get('dtstart').dt
        dtend = component.get('dtend').dt

        # All-day detection (standard all-day or 00:00 start with full-day multiple)
        is_all_day = not isinstance(dtstart, datetime.datetime)
        if not is_all_day and isinstance(dtstart, datetime.datetime):
            starts_at_midnight = dtstart.time() == datetime.time(0, 0)
            duration = dtend - dtstart
            is_full_day_multiple = duration.total_seconds() > 0 and duration.total_seconds() % (24 * 3600) == 0
            if starts_at_midnight and is_full_day_multiple:
                is_all_day = True

        if is_all_day:
            start_date = dtstart if not isinstance(dtstart, datetime.datetime) else dtstart.date()
            end_date = dtend if not isinstance(dtend, datetime.datetime) else dtend.date()
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

    # ---------- Sync core (now with optional batching) ----------

    def sync_events(self):
        calendar_id = self.config.get('google_calendar_id')
        days_past = self.config.get('sync_days_past', 90)
        days_future = self.config.get('sync_days_future', 120)

        now = self.tz.localize(datetime.datetime.now())
        time_min = now - datetime.timedelta(days=days_past)
        time_max = now + datetime.timedelta(days=days_future)

        if not os.path.exists(self.export_file):
            self.logger.error(f"Export file not found at {self.export_file}. Aborting sync.")
            return

        expanded_events = self.parse_ics_file(self.export_file, time_min, time_max)
        existing_google = self.get_existing_google_events(calendar_id, time_min, time_max)

        # Build operation lists
        to_create = []
        to_update = []
        outlook_uids = set()
        processed_uids = set()

        for component in expanded_events:
            try:
                uid = self.compute_uid(component)
                outlook_uids.add(uid)
                if uid in processed_uids:
                    continue
                body = self.build_event_body(component, uid)
                if uid not in existing_google:
                    to_create.append(('insert', body))
                else:
                    existing_event = existing_google[uid]
                    changed = False
                    for key in ('summary', 'location', 'description'):
                        if body.get(key, '') != existing_event.get(key, ''):
                            changed = True
                            break
                    if changed:
                        to_update.append(('update', existing_event['id'], body))
                processed_uids.add(uid)
            except Exception as e:
                self.logger.error(f"Unexpected error processing event '{component.get('summary','Unknown Event')}': {e}", exc_info=True)

        # Orphans to delete
        orphan_uids = set(existing_google.keys()) - outlook_uids
        to_delete = []
        for uid in orphan_uids:
            ev = existing_google[uid]
            to_delete.append(('delete', ev['id'], ev.get('summary', 'No Title')))

        # Execute operations: batch if enabled, else one by one (old behavior)
        if self.enable_batch:
            self.logger.info(f"Batch mode ON. create={len(to_create)}, update={len(to_update)}, delete={len(to_delete)}")
            self._run_batched(calendar_id, to_create, to_update, to_delete)
        else:
            self.logger.info(f"Batch mode OFF. create={len(to_create)}, update={len(to_update)}, delete={len(to_delete)}")
            # one-by-one (preserves previous behavior)
            for _, body in to_create:
                self.logger.info(f"Creating: {body['summary']}")
                req = self.google_service.events().insert(calendarId=calendar_id, body=body)
                self._exec_with_backoff(req)
                time.sleep(self.api_delay)
            for _, event_id, body in to_update:
                self.logger.info(f"Updating: {body['summary']}")
                req = self.google_service.events().update(calendarId=calendar_id, eventId=event_id, body=body)
                self._exec_with_backoff(req)
                time.sleep(self.api_delay)
            for _, event_id, summary in to_delete:
                self.logger.info(f"Deleting orphan: {summary}")
                req = self.google_service.events().delete(calendarId=calendar_id, eventId=event_id)
                self._exec_with_backoff(req)
                time.sleep(self.api_delay)

        # Clean up export file
        try:
            if os.path.exists(self.export_file):
                os.remove(self.export_file)
        except Exception:
            pass

    # ---------- Batch helpers ----------

    def _run_batched(self, calendar_id, to_create, to_update, to_delete):
        """
        Use googleapiclient's batch to group operations.
        Each batch sends up to self.batch_size requests; we run create, update, delete in that order.
        Failures are logged per-item and do not stop the batch.
        """
        def run_batch_group(reqs):
            # Build batch and execute
            batch = self.google_service.new_batch_http_request()
            for idx, (kind, payload) in enumerate(reqs):
                def _callback(kind=kind, payload=payload):
                    def cb(request_id, response, exception):
                        if exception:
                            # Log but continue
                            if kind == 'insert':
                                self.logger.error(f"[batch] Create FAILED: {payload.get('summary','(no summary)')} :: {exception}")
                            elif kind == 'update':
                                self.logger.error(f"[batch] Update FAILED: {payload.get('summary','(no summary)')} :: {exception}")
                            elif kind == 'delete':
                                self.logger.error(f"[batch] Delete FAILED: {payload} :: {exception}")
                        else:
                            # Optional success logging at debug level to keep output tidy
                            pass
                    return cb

                if kind == 'insert':
                    req = self.google_service.events().insert(calendarId=calendar_id, body=payload)
                elif kind == 'update':
                    event_id, body = payload  # payload is tuple here
                    req = self.google_service.events().update(calendarId=calendar_id, eventId=event_id, body=body)
                elif kind == 'delete':
                    event_id = payload
                    req = self.google_service.events().delete(calendarId=calendar_id, eventId=event_id)
                else:
                    continue

                batch.add(req, callback=_callback())

            # Execute with top-level retry in case the whole batch fails transiently
            self._exec_batch_with_backoff(batch)

        # Flatten payloads into a uniform list for batching: ('insert', body) -> ('insert', body)
        # For updates/deletes we wrap to make callback logging simple.
        create_payloads = [('insert', body) for _, body in to_create]
        update_payloads = [('update', (event_id, body)) for _, event_id, body in to_update]
        delete_payloads = [('delete', event_id) for _, event_id, _ in to_delete]

        # Chunk and send
        def chunks(arr, n):
            for i in range(0, len(arr), n):
                yield arr[i:i+n]

        for group_name, items in (('create', create_payloads), ('update', update_payloads), ('delete', delete_payloads)):
            if not items:
                continue
            self.logger.info(f"Executing {group_name} in batches of {self.batch_size} (total {len(items)})")
            for chunk in chunks(items, self.batch_size):
                run_batch_group(chunk)
                # Light delay between batches to be kind to quota
                time.sleep(max(self.api_delay, 0.05))

    def _exec_with_backoff(self, api_request):
        max_retries = 5
        for attempt in range(max_retries):
            try:
                return api_request.execute()
            except HttpError as e:
                status = e.resp.status
                if status == 409:
                    self.logger.warning("Conflict (likely duplicate). Ignoring.")
                    return None
                if status in [403, 500, 503]:
                    if attempt < max_retries - 1:
                        delay = (2 ** attempt) + random.uniform(0, 1)
                        self.logger.warning(f"API rate limit/server error. Retrying in {delay:.2f}s...")
                        time.sleep(delay)
                    else:
                        self.logger.error("Max retries exceeded.")
                        raise
                elif status in [404, 410]:
                    self.logger.warning("Resource not found or already deleted.")
                    return None
                else:
                    raise

    def _exec_batch_with_backoff(self, batch):
        max_retries = 5
        for attempt in range(max_retries):
            try:
                batch.execute()
                return
            except HttpError as e:
                status = e.resp.status if hasattr(e, 'resp') else None
                if status in [403, 500, 503] and attempt < max_retries - 1:
                    delay = (2 ** attempt) + random.uniform(0, 1)
                    self.logger.warning(f"[batch] API rate limit/server error. Retrying batch in {delay:.2f}s...")
                    time.sleep(delay)
                else:
                    # If the whole batch bombs out, log and move on;
                    # per-item operations still survive in subsequent batches.
                    self.logger.error(f"[batch] Batch failed irrecoverably: {e}")
                    return

    # ---------- driver ----------

    def run(self):
        self.logger.info("--- Starting CalendarBridge Sync Cycle ---")
        if self.run_applescript_export():
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
    CalendarSync().run()
