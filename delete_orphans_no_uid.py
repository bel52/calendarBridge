#!/usr/bin/env python3
"""
delete_orphans_no_uid.py  –  remove Google events that
   • are in the rolling window
   • LACK extendedProperties.private.icalUID
   • have the same summary & start as an Outlook event
     (meaning they’re true duplicates)
"""

import os, glob, logging, time
from datetime import datetime, timedelta, timezone, date
from icalendar import Calendar
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

WIN_BACK, WIN_AHEAD = 7, 120
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
            raise RuntimeError("Google creds bad")
    return build('calendar', 'v3', credentials=creds, cache_discovery=False)

def load_outlook_index(start, end):
    """Return {(summary,start_iso): True} set for quick lookup."""
    idx = set()
    for fp in glob.glob(os.path.join(ICS_DIR, '*.ics')):
        with open(fp, 'rb') as f:
            cal = Calendar.from_ical(f.read())
        for c in cal.walk():
            if c.name != 'VEVENT': continue
            dt = utc_naive(c.decoded('DTSTART'))
            if start <= dt <= end:
                key = (str(c.get('SUMMARY','')).strip(), dt.isoformat())
                idx.add(key)
    return idx

def main():
    win_start, win_end = window_limits()
    logging.info(f"Dedup window {win_start.date()} → {win_end.date()}")

    outlook_idx = load_outlook_index(win_start, win_end)
    logging.info(f"Outlook events in window: {len(outlook_idx)}")

    svc, page, deleted = google_service(), None, 0
    while True:
        resp = svc.events().list(calendarId=CAL_ID, singleEvents=True, showDeleted=False,
                                 timeMin=win_start.isoformat()+'Z',
                                 timeMax=win_end.isoformat()+'Z',
                                 pageToken=page).execute()
        for ev in resp.get('items', []):
            uid = ev.get('extendedProperties', {}).get('private', {}).get('icalUID')
            if uid:  # leave those for the other cleaner
                continue
            summary = ev.get('summary','').strip()
            start_raw = ev['start'].get('dateTime', ev['start'].get('date'))
            dt = utc_naive(datetime.fromisoformat(start_raw))
            if (summary, dt.isoformat()) in outlook_idx:
                try:
                    svc.events().delete(calendarId=CAL_ID, eventId=ev['id']).execute()
                    deleted += 1
                    logging.info(f"Deleted dup w/o UID: {summary} @ {dt}")
                except Exception as e:
                    logging.warning(f"Fail delete {ev['id']}: {e}")
            if deleted % MAX_RPS == 0:
                time.sleep(1)
        page = resp.get('nextPageToken')
        if not page: break

    logging.info(f"✅ Deleted duplicates without UID: {deleted}")

if __name__ == '__main__':
    main()
