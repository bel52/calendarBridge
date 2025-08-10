#!/usr/bin/env python3
"""
delete_orphans.py – remove Google Calendar events that:
  • have extendedProperties.private.icalUID
  • UID does NOT exist in Outlook’s current 7-day-back ⇢ 120-day-ahead window
"""

import os, glob, logging, time
from datetime import datetime, timedelta, timezone, date
from icalendar import Calendar
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.auth.transport.requests import Request

# Window must match safe_sync.py
WIN_BACK  = 7
WIN_AHEAD = 120

SCOPES     = ['https://www.googleapis.com/auth/calendar']
TOKEN_FILE = os.path.expanduser('~/calendarBridge/token.json')
ICS_DIR    = os.path.expanduser('~/calendarBridge/outbox')
CAL_ID     = 'primary'
MAX_RPS    = 8

logging.basicConfig(level=logging.INFO, format='%(levelname)s %(message)s')

def utc_naive(dt):
    if isinstance(dt, datetime):
        return dt.astimezone(timezone.utc).replace(tzinfo=None)
    else:
        return datetime.combine(dt, datetime.min.time())

def window_limits():
    now = datetime.utcnow().replace(tzinfo=None)
    return now - timedelta(days=WIN_BACK), now + timedelta(days=WIN_AHEAD)

def google_service():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise RuntimeError("Google creds invalid")
    return build('calendar', 'v3', credentials=creds, cache_discovery=False)

def load_outlook_uids(start, end):
    uids = set()
    for fp in glob.glob(os.path.join(ICS_DIR, '*.ics')):
        with open(fp, 'rb') as f:
            cal = Calendar.from_ical(f.read())
        for c in cal.walk():
            if c.name != 'VEVENT':
                continue
            uid = str(c.get('UID'))
            dt  = utc_naive(c.decoded('DTSTART'))
            if start <= dt <= end:
                uids.add(uid)
    return uids

def main():
    win_start, win_end = window_limits()
    logging.info(f"Cleanup window: {win_start.date()} → {win_end.date()}")

    outlook_uids = load_outlook_uids(win_start, win_end)
    logging.info(f"Outlook UIDs in window: {len(outlook_uids)}")

    svc = google_service()
    page, deleted = None, 0
    while True:
        resp = svc.events().list(calendarId=CAL_ID, singleEvents=True, showDeleted=False,
                                 timeMin=win_start.isoformat()+'Z',
                                 timeMax=win_end.isoformat()+'Z',
                                 pageToken=page).execute()
        for ev in resp.get('items', []):
            uid = ev.get('extendedProperties', {}).get('private', {}).get('icalUID')
            if uid and uid not in outlook_uids:
                try:
                    svc.events().delete(calendarId=CAL_ID, eventId=ev['id']).execute()
                    deleted += 1
                    logging.info(f"Deleted orphan: {ev.get('summary','')}")
                except HttpError as e:
                    logging.warning(f"Failed delete id={ev['id']}: {e}")
            if deleted % MAX_RPS == 0:
                time.sleep(1)
        page = resp.get('nextPageToken')
        if not page: break

    logging.info(f"✅ Orphans deleted: {deleted}")

if __name__ == '__main__':
    main()
