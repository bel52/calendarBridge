#!/usr/bin/env python3
import os, sys, datetime, re
from zoneinfo import ZoneInfo
from collections import defaultdict
from icalendar import Calendar
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request

BASE   = os.path.expanduser('~/calendarBridge')
ICS_DIR = os.path.join(BASE, 'outbox')
CAL_ID  = 'primary'
DEFAULT_TZ_NAME = os.environ.get('CALBRIDGE_TZ', 'America/New_York')
try:
    DEFAULT_TZ = ZoneInfo(DEFAULT_TZ_NAME)
except Exception:
    DEFAULT_TZ_NAME, DEFAULT_TZ = 'UTC', ZoneInfo('UTC')

def ensure_tz(dt):
    if isinstance(dt, datetime.datetime) and (dt.tzinfo is None or dt.tzinfo.utcoffset(dt) is None):
        return dt.replace(tzinfo=DEFAULT_TZ)
    return dt

def canonical_recur_id(rec_prop) -> str:
    if not rec_prop: return ''
    v = getattr(rec_prop, 'dt', rec_prop)
    if isinstance(v, datetime.date) and not isinstance(v, datetime.datetime):
        return v.strftime('%Y%m%d')
    dt = ensure_tz(v).astimezone(datetime.timezone.utc)
    return dt.strftime('%Y%m%dT%H%M%SZ')

def exdate_for_base(dt_like, base_all_day: bool) -> str:
    if base_all_day:
        if isinstance(dt_like, datetime.datetime):
            dt_like = dt_like.date()
        return "EXDATE:" + dt_like.strftime("%Y%m%d")
    if isinstance(dt_like, datetime.date) and not isinstance(dt_like, datetime.datetime):
        dt_like = datetime.datetime(dt_like.year, dt_like.month, dt_like.day, tzinfo=DEFAULT_TZ)
    local = ensure_tz(dt_like).astimezone(DEFAULT_TZ)
    return f"EXDATE;TZID={DEFAULT_TZ_NAME}:" + local.strftime("%Y%m%dT%H%M%S")

def auth_readonly():
    SCOPES = ['https://www.googleapis.com/auth/calendar.readonly']
    token = os.path.join(BASE,'token.json')
    credf = os.path.join(BASE,'credentials.json')
    creds = None
    if os.path.exists(token):
        creds = Credentials.from_authorized_user_file(token, SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            creds = InstalledAppFlow.from_client_secrets_file(credf, SCOPES).run_local_server(port=0)
        with open(token,'w') as f: f.write(creds.to_json())
    return build('calendar','v3',credentials=creds, cache_discovery=False)

def load_google_map(service):
    google = {}
    page = None
    while True:
        resp = service.events().list(calendarId=CAL_ID, pageToken=page, maxResults=2500,
                                     fields='nextPageToken,items(id,summary,start,end,extendedProperties)').execute()
        for it in resp.get('items', []):
            key = it.get('extendedProperties', {}).get('private', {}).get('compositeUID')
            if key:
                google[key] = it
        page = resp.get('nextPageToken')
        if not page: break
    return google

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python diag_series.py 'partial title text'")
        sys.exit(1)
    needle = sys.argv[1].lower()

    # parse Outlook ICS
    outlook = {}
    series_names = defaultdict(set)
    for fn in os.listdir(ICS_DIR):
        if not fn.endswith('.ics'): continue
        cal = Calendar.from_ical(open(os.path.join(ICS_DIR,fn),'rb').read())
        for ev in cal.walk('VEVENT'):
            title = str(ev.get('SUMMARY','')).lower()
            if needle not in title: continue
            uid = str(ev.get('UID'))
            rec = canonical_recur_id(ev.get('RECURRENCE-ID'))
            outlook[f"{uid}❄{rec}"] = ev
            series_names[uid].add(str(ev.get('SUMMARY','')))

    if not outlook:
        print("No matching Outlook events found in outbox/. Did you run export & clean?")
        sys.exit(0)

    print(f"Found series UIDs:")
    for uid,names in series_names.items():
        print(f"  UID {uid}  titles={list(names)[:1]}{' …' if len(names)>1 else ''}")

    svc = auth_readonly()
    gmap = load_google_map(svc)

    for uid in series_names.keys():
        base_key = f"{uid}❄"
        base = outlook.get(base_key)
        if not base:
            print(f"\nUID {uid}: (no base master in window; only exceptions present)")
            base_is_all_day = False
        else:
            s = base['DTSTART'].dt
            e = base.get('DTEND', base['DTSTART']).dt
            base_is_all_day = isinstance(s, datetime.date) and not isinstance(s, datetime.datetime)

        print(f"\n=== UID {uid} ===")
        print(f"Google base present? {'YES' if base_key in gmap else 'NO'}")

        # exceptions seen in Outlook
        exc_keys = sorted([k for k in outlook.keys() if k.startswith(uid+'❄') and k != base_key])
        if not exc_keys:
            print("No modified exceptions found.")
        else:
            print("Modified exceptions (canonical keys) and EXDATE that base will carry:")
            for k in exc_keys:
                rec_prop = outlook[k].get('RECURRENCE-ID')
                dtlike = getattr(rec_prop,'dt',rec_prop)
                print(f"  {k}   EXDATE={exdate_for_base(dtlike, base_is_all_day)}")

        # what Google has for this series
        print("Google items for this series (by compositeUID):")
        for gkey, item in gmap.items():
            if gkey.startswith(uid+'❄'):
                st = item['start'].get('dateTime') or (item['start'].get('date') + ' (all-day)')
                print(f"  {gkey}   id={item['id']}   start={st}")
