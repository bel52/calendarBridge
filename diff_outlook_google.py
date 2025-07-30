#!/usr/bin/env python3
"""
diff_outlook_google.py – compare Outlook-exported events to Google Calendar.

• Saves CSV report to ~/calendarBridge/logs/diff_report.csv
• Prints counts of:
      missing_on_google   → Outlook UID not found on Google
      orphan_on_google    → Google event has icalUID but no Outlook match
      mismatch_body/time  → UID exists both places but body differs
"""

import os, csv, glob, logging
from datetime import datetime, timedelta, date, timezone
from icalendar import Calendar
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

# ——— Config (matches safe_sync.py) ————————————————————————
WIN_BACK   = 7
WIN_AHEAD  = 120
SCOPES     = ['https://www.googleapis.com/auth/calendar']   # full scope
TOKEN_FILE = os.path.expanduser('~/calendarBridge/token.json')
ICS_DIR    = os.path.expanduser('~/calendarBridge/outbox')
CAL_ID     = 'primary'
LOG_DIR    = os.path.expanduser('~/calendarBridge/logs')
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(level=logging.INFO, format='%(levelname)s %(message)s')

def utc_naive(dt):
    if isinstance(dt, datetime):
        return dt.astimezone(timezone.utc).replace(tzinfo=None)
    else:
        return datetime.combine(dt, datetime.min.time())

# ——— Load Outlook ————————————————————————————
def load_outlook(win_start, win_end):
    outlook = {}
    for fp in glob.glob(os.path.join(ICS_DIR, '*.ics')):
        try:
            with open(fp, 'rb') as f:
                cal = Calendar.from_ical(f.read())
            for comp in cal.walk():
                if comp.name != 'VEVENT':
                    continue
                uid = str(comp.get('UID'))
                dt  = utc_naive(comp.decoded('DTSTART'))
                if win_start <= dt <= win_end:
                    outlook[uid] = {
                        'uid'     : uid,
                        'summary' : str(comp.get('SUMMARY', '')),
                        'start'   : dt,
                    }
        except Exception as e:
            logging.warning(f"Bad ICS {fp}: {e}")
    return outlook

# ——— Load Google —————————————————————————————
def google_service():
    creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES)
    if not creds.valid:
        if creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            raise RuntimeError("Google credentials invalid")
    return build('calendar', 'v3', credentials=creds, cache_discovery=False)

def load_google(win_start, win_end):
    svc = google_service()
    g_events, page = {}, None
    while True:
        resp = svc.events().list(calendarId=CAL_ID, singleEvents=True, showDeleted=False,
                                 timeMin=win_start.isoformat()+'Z',
                                 timeMax=win_end.isoformat()+'Z',
                                 pageToken=page).execute()
        for item in resp.get('items', []):
            uid = item.get('extendedProperties', {}).get('private', {}).get('icalUID')
            key = uid if uid else f"NOUID-{item['id']}"
            start_raw = item['start'].get('dateTime', item['start'].get('date'))
            g_events[key] = {
                'uid'     : uid if uid else '',
                'id'      : item['id'],
                'summary' : item.get('summary', ''),
                'start'   : utc_naive(datetime.fromisoformat(start_raw)),
            }
        page = resp.get('nextPageToken')
        if not page: return g_events

# ——— Main diff ———————————————————————————————
def main():
    win_start = datetime.utcnow().replace(tzinfo=None) - timedelta(days=WIN_BACK)
    win_end   = datetime.utcnow().replace(tzinfo=None) + timedelta(days=WIN_AHEAD)
    logging.info(f"Diff window: {win_start.date()} → {win_end.date()}")

    outlook = load_outlook(win_start, win_end)
    google  = load_google(win_start, win_end)

    missing, orphan, mismatch = [], [], []

    for uid, ev in outlook.items():
        if uid not in google:
            missing.append(ev)

    for key, ev in google.items():
        uid = ev['uid']
        if not uid:
            orphan.append(ev)
        elif uid in outlook:
            out_ev = outlook[uid]
            if abs((out_ev['start'] - ev['start']).total_seconds()) > 60 \
               or out_ev['summary'] != ev['summary']:
                mismatch.append({'uid': uid, 'g_id': ev['id'],
                                 'g_summary': ev['summary'],
                                 'o_summary': out_ev['summary']})
        else:
            orphan.append(ev)          # UID not present in Outlook window

    # CSV output
    csv_path = os.path.join(LOG_DIR, 'diff_report.csv')
    with open(csv_path, 'w', newline='') as f:
        w = csv.writer(f)
        w.writerow(['type','uid','google_id','summary','start'])
        for ev in missing:
            w.writerow(['missing', ev['uid'],'', ev['summary'], ev['start']])
        for ev in orphan:
            w.writerow(['orphan', ev.get('uid',''), ev.get('id',''),
                        ev['summary'], ev['start']])
        for ev in mismatch:
            w.writerow(['mismatch', ev['uid'], ev['g_id'],
                        f"G:{ev['g_summary']} / O:{ev['o_summary']}", ''])

    logging.info(f"Report saved: {csv_path}")
    logging.info(f"Missing on Google  : {len(missing)}")
    logging.info(f"Orphan duplicates  : {len(orphan)}")
    logging.info(f"Body mismatches    : {len(mismatch)}")

if __name__ == '__main__':
    main()
