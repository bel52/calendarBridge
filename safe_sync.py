#!/usr/bin/env python3
"""
safe_sync.py  –  Outlook-to-Google incremental sync

2025-07-30  v2.2
• Window-aware export:  7 days back → 120 days ahead
• Throttles to ≤ 8 Google requests/sec
• Exponential back-off on BOTH:
      – 403 rateLimitExceeded
      – 500 backendError
• Skips out-of-window events to avoid needless processing
"""

import os, glob, logging, time
from datetime import datetime, timedelta, date, timezone
from icalendar import Calendar
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request

# ─── Tunables ──────────────────────────────────────────────────
WINDOW_BACK_DAYS   = 7
WINDOW_AHEAD_DAYS  = 120
BATCH_INSERT_LIMIT = 3000   # safety cap per run
MAX_RPS            = 8      # requests per second (stay < Google 10)
MAX_RETRIES        = 5
RETRY_BACKOFF_BASE = 2      # exponential factor 2,4,8,16,32 s

SCOPES     = ['https://www.googleapis.com/auth/calendar']
TOKEN_FILE = os.path.expanduser('~/calendarBridge/token.json')
ICS_DIR    = os.path.expanduser('~/calendarBridge/outbox')
CAL_ID     = 'primary'

logging.basicConfig(level=logging.INFO,
                    format='%(levelname)s %(message)s')

# ─── Helper functions ──────────────────────────────────────────
def utc_naive(dt):
    """Return a timezone-naïve UTC datetime (or 00:00 for VALUE=DATE)."""
    if isinstance(dt, datetime):
        return dt.astimezone(timezone.utc).replace(tzinfo=None)
    else:  # all-day date
        return datetime.combine(dt, datetime.min.time())

def window_limits():
    now = datetime.utcnow().replace(tzinfo=None)
    return (now - timedelta(days=WINDOW_BACK_DAYS),
            now + timedelta(days=WINDOW_AHEAD_DAYS))

def google_service():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise RuntimeError("Google credentials invalid or missing.")
    return build('calendar', 'v3', credentials=creds, cache_discovery=False)

def load_local_events():
    """Return {uid: (component, dt_start_utc_naive)} for all .ics files."""
    events = {}
    for fp in glob.glob(os.path.join(ICS_DIR, '*.ics')):
        try:
            with open(fp, 'rb') as f:
                cal = Calendar.from_ical(f.read())
            for comp in cal.walk():
                if comp.name == 'VEVENT':
                    uid = str(comp.get('UID'))
                    dtstart = utc_naive(comp.decoded('DTSTART'))
                    events[uid] = (comp, dtstart)
        except Exception as e:
            logging.warning(f"Skip bad ICS {fp}: {e}")
    return events

def fetch_remote_uids(svc, win_start, win_end):
    """Return {icalUID: google_event_id} for already-synced events in window."""
    uids, page = {}, None
    while True:
        resp = svc.events().list(calendarId=CAL_ID,
                                 timeMin=win_start.isoformat()+'Z',
                                 timeMax=win_end.isoformat()+'Z',
                                 singleEvents=True,
                                 showDeleted=False,
                                 pageToken=page).execute()
        for item in resp.get('items', []):
            uid = item.get('extendedProperties', {}).get('private', {}).get('icalUID')
            if uid:
                uids[uid] = item['id']
        page = resp.get('nextPageToken')
        if not page:
            return uids

def build_body(comp):
    uid   = str(comp.get('UID'))
    start = comp.decoded('DTSTART')
    end   = comp.decoded('DTEND')
    allday = isinstance(start, date) and not isinstance(start, datetime)

    body = {
        'summary':      str(comp.get('SUMMARY', '')),
        'location':     str(comp.get('LOCATION', '')),
        'description':  str(comp.get('DESCRIPTION', '')),
        'extendedProperties': {'private': {'icalUID': uid}},
    }
    if allday:
        body['start'] = {'date': start.isoformat()}
        body['end']   = {'date': end.isoformat()}
    else:
        body['start'] = {'dateTime': start.astimezone(timezone.utc).isoformat(),
                         'timeZone': 'UTC'}
        body['end']   = {'dateTime': end.astimezone(timezone.utc).isoformat(),
                         'timeZone': 'UTC'}
    return body

def call_with_backoff(call):
    """Execute a Google API request with retry/back-off on quota or 500 errors."""
    for attempt in range(MAX_RETRIES):
        try:
            return call.execute()
        except HttpError as e:
            msg = str(e)
            if (
                (e.resp.status == 403 and 'rateLimitExceeded' in msg)
                or (e.resp.status == 500 and 'backendError'   in msg)
            ):
                wait = RETRY_BACKOFF_BASE ** attempt
                logging.warning(f"{e.resp.status} – backing off {wait}s …")
                time.sleep(wait)
            else:
                raise
    logging.error("Exceeded retries; skipping this event.")
    return None

# ─── Main sync routine ─────────────────────────────────────────
def main():
    win_start, win_end = window_limits()
    logging.info(f"Window: {win_start.date()} → {win_end.date()}")

    local  = load_local_events()
    logging.info(f"📂 Parsed {len(local)} local events")

    svc    = google_service()
    remote = fetch_remote_uids(svc, win_start, win_end)
    logging.info(f"☁️  Remote Outlook-origin events: {len(remote)}")

    to_ins, to_upd, skipped = [], [], 0
    for uid, (comp, dt) in local.items():
        if not (win_start <= dt <= win_end):
            skipped += 1
            continue
        body = build_body(comp)
        if uid in remote:
            to_upd.append((remote[uid], body))
        else:
            to_ins.append(body)

    # safety cap
    to_ins = to_ins[:BATCH_INSERT_LIMIT]

    ins = upd = 0
    ops = [('insert', None, b) for b in to_ins] + \
          [('update', eid, b) for eid, b in to_upd]

    for i, (kind, eid, body) in enumerate(ops, 1):
        if kind == 'insert':
            if call_with_backoff(svc.events().insert(calendarId=CAL_ID, body=body)):
                ins += 1
        else:
            if call_with_backoff(svc.events().update(calendarId=CAL_ID, eventId=eid, body=body)):
                upd += 1

        if i % MAX_RPS == 0:
            time.sleep(1)

    logging.info(f"✅ Sync ➕{ins} 🔄{upd} ⏭{skipped}")

if __name__ == '__main__':
    main()
