#!/usr/bin/env python3
"""
safe_sync.py  –  Outlook ➜ Google Calendar one-way sync

• 60-second global socket timeout  
• Automatic 3× retry with 5-s back-off on Calendar API calls  
• Detects midnight-to-midnight spans and sends them as true all-day events
"""
import os, json, time, socket, hashlib, datetime
from collections import deque

# ───────────────────── socket hard-timeout ───────────────────────────────
socket.setdefaulttimeout(60)

def retry(fn, *a, **kw):
    """Run fn() with retries on socket timeout or 403/429 Google rate-limits."""
    for attempt in range(3):
        try:
            return fn(*a, **kw)
        except (socket.timeout, TimeoutError):
            if attempt == 2:
                raise
            time.sleep(5)
        except Exception as e:
            if hasattr(e, 'resp') and getattr(e.resp, 'status', 0) in (403, 429):
                time.sleep(5)
            else:
                raise

# ───────────────────── Google auth / build ───────────────────────────────
from google.oauth2.credentials  import Credentials
from google_auth_oauthlib.flow  import InstalledAppFlow
from googleapiclient.discovery  import build
from googleapiclient.errors     import HttpError
from google.auth.transport.requests import Request

BASE   = os.path.expanduser('~/calendarBridge')
TOKEN  = f"{BASE}/token.json"
CREDS  = f"{BASE}/credentials.json"
SCOPES = ['https://www.googleapis.com/auth/calendar']
CAL_ID = 'primary'

creds = None
if os.path.exists(TOKEN):
    creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        creds = InstalledAppFlow.from_client_secrets_file(CREDS, SCOPES).run_local_server(port=0)
    with open(TOKEN, 'w') as f:
        f.write(creds.to_json())

service = build('calendar', 'v3', credentials=creds, cache_discovery=False)

# ───────────────────── read Outlook .ics files ───────────────────────────
from icalendar import Calendar
ICS_DIR = f"{BASE}/outbox"
outlook = {}
for fn in os.listdir(ICS_DIR):
    if fn.endswith('.ics'):
        with open(f"{ICS_DIR}/{fn}", 'rb') as f:
            cal = Calendar.from_ical(f.read())
            for ev in cal.walk('VEVENT'):
                uid = str(ev.get('UID'))
                rec = str(ev.get('RECURRENCE-ID', ''))
                outlook[f"{uid}❄{rec}"] = ev
print(f"📂 Parsed {len(outlook)} Outlook events")

# Remove canceled recurring occurrences from outlook (so they will be deleted)
for key in list(outlook.keys()):
    comp = outlook[key]
    if str(comp.get('STATUS', '')).upper() == 'CANCELLED':
        outlook.pop(key, None)

# ───────────────────── pull Google events we manage ──────────────────────
google, page = {}, None
while True:
    resp = retry(service.events().list,
                 calendarId=CAL_ID,
                 pageToken=page,
                 maxResults=2500,
                 fields=('nextPageToken, items(id,summary,description,start,end,extendedProperties)')).execute()
    for it in resp.get('items', []):
        key = it.get('extendedProperties', {}).get('private', {}).get('compositeUID')
        if key:
            google[key] = it
    page = resp.get('nextPageToken')
    if not page:
        break
print(f"☁️  Loaded {len(google)} Google events tagged as Outlook-origin")

# ───────────────────── helpers ───────────────────────────────────────────
def to_gcal_dt(dt):
    return {'dateTime': dt.astimezone().isoformat()} if isinstance(dt, datetime.datetime) else {'date': dt.isoformat()}

def body_hash(b):
    return hashlib.sha256(json.dumps(b, sort_keys=True).encode()).hexdigest()

def maybe_all_day(start, end):
    '''If span is midnight-to-midnight whole-days, return (d1,d2) else None.'''
    if isinstance(start, datetime.datetime) and isinstance(end, datetime.datetime):
        if start.time() == end.time() == datetime.time(0, 0) and (end - start).seconds == 0:
            return start.date(), end.date()
    return None

DELETE_BATCH = 50
QUAR_FILE    = f"{BASE}/quarantine.txt"
quarantined  = set(open(QUAR_FILE).read().split()) if os.path.exists(QUAR_FILE) else set()

# ───────────────────── add / update phase ────────────────────────────────
added = updated = skipped = 0

# Group events by base UID (to handle recurring events with exceptions)
grouped = {}
for comp_key, comp in outlook.items():
    base_uid = comp_key.split('❄')[0]
    grouped.setdefault(base_uid, []).append(comp_key)

for base_uid, comp_keys in grouped.items():
    base_comp_key = base_uid + '❄'
    base_comp = outlook.get(base_comp_key)
    if base_comp is None:
        continue
    if base_uid in quarantined:
        continue

    if base_comp.get('RRULE'):
        # Recurring event with possible exceptions
        s = base_comp['DTSTART'].dt
        e = base_comp.get('DTEND', base_comp['DTSTART']).dt
        ad = maybe_all_day(s, e)
        if ad:
            g_start = {'date': ad[0].isoformat()}
            g_end   = {'date': ad[1].isoformat()}
        else:
            if isinstance(s, datetime.date) and not isinstance(s, datetime.datetime):
                e = e if isinstance(e, datetime.date) else s + datetime.timedelta(days=1)
            g_start, g_end = to_gcal_dt(s), to_gcal_dt(e)
        body = {
            'summary':  str(base_comp.get('SUMMARY', 'No title')),
            'location': str(base_comp.get('LOCATION', '')),
            'start':    g_start,
            'end':      g_end,
            'extendedProperties': {'private': {'compositeUID': base_comp_key}}
        }
        # Build recurrence rules/exceptions
        recurrences = []
        rrule_prop = base_comp.get('RRULE')
        if rrule_prop:
            rrule_str = rrule_prop.to_ical().decode() if hasattr(rrule_prop, 'to_ical') else str(rrule_prop)
            rrule_str = rrule_str.strip()
            if not rrule_str.upper().startswith('RRULE'):
                rrule_str = 'RRULE:' + rrule_str
            recurrences.append(rrule_str)
        if base_comp.get('RDATE'):
            rdate_prop = base_comp.get('RDATE')
            rdate_str = rdate_prop.to_ical().decode() if hasattr(rdate_prop, 'to_ical') else str(rdate_prop)
            for ln in rdate_str.splitlines():
                ln = ln.strip()
                if not ln:
                    continue
                if not ln.upper().startswith('RDATE'):
                    ln = 'RDATE:' + ln
                recurrences.append(ln)
        if base_comp.get('EXDATE'):
            exdate_prop = base_comp.get('EXDATE')
            exdate_str = exdate_prop.to_ical().decode() if hasattr(exdate_prop, 'to_ical') else str(exdate_prop)
            for ln in exdate_str.splitlines():
                ln = ln.strip()
                if not ln:
                    continue
                if not ln.upper().startswith('EXDATE'):
                    ln = 'EXDATE:' + ln
                recurrences.append(ln)
        # Include exceptions (for exclusions and overrides)
        for comp_key in comp_keys:
            if comp_key == base_comp_key:
                continue
            exc_comp = outlook[comp_key]
            recurid_prop = exc_comp.get('RECURRENCE-ID')
            if recurid_prop:
                recurid_value = recurid_prop.to_ical().decode().strip() if hasattr(recurid_prop, 'to_ical') else str(recurid_prop)
                exdate_line = 'EXDATE'
                for param, val in recurid_prop.params.items():
                    exdate_line += f';{param}={val}'
                exdate_line += ':' + recurid_value
                if exdate_line not in recurrences:
                    recurrences.append(exdate_line)
            status = str(exc_comp.get('STATUS', ''))
            if status.upper() == 'CANCELLED':
                # No separate event for cancelled occurrence (just excluded above)
                continue
            # Add modified occurrence as separate event
            s_exc = exc_comp['DTSTART'].dt
            e_exc = exc_comp.get('DTEND', exc_comp['DTSTART']).dt
            ad_exc = maybe_all_day(s_exc, e_exc)
            if ad_exc:
                g_start_exc = {'date': ad_exc[0].isoformat()}
                g_end_exc   = {'date': ad_exc[1].isoformat()}
            else:
                if isinstance(s_exc, datetime.date) and not isinstance(s_exc, datetime.datetime):
                    e_exc = e_exc if isinstance(e_exc, datetime.date) else s_exc + datetime.timedelta(days=1)
                g_start_exc, g_end_exc = to_gcal_dt(s_exc), to_gcal_dt(e_exc)
            exc_body = {
                'summary':  str(exc_comp.get('SUMMARY', base_comp.get('SUMMARY', 'No title'))),
                'location': str(exc_comp.get('LOCATION', base_comp.get('LOCATION', ''))),
                'start':    g_start_exc,
                'end':      g_end_exc,
                'extendedProperties': {'private': {'compositeUID': comp_key}}
            }
            exc_body['description'] = body_hash(exc_body)
            g_evt_exc = google.get(comp_key)
            if not g_evt_exc:
                retry(service.events().insert, calendarId=CAL_ID, body=exc_body).execute()
                added += 1
            elif g_evt_exc.get('description') != exc_body['description']:
                retry(service.events().update, calendarId=CAL_ID, eventId=g_evt_exc['id'], body=exc_body).execute()
                updated += 1
            else:
                skipped += 1
        if recurrences:
            body['recurrence'] = recurrences
        body['description'] = body_hash(body)
        g_evt_base = google.get(base_comp_key)
        if not g_evt_base:
            retry(service.events().insert, calendarId=CAL_ID, body=body).execute()
            added += 1
        elif g_evt_base.get('description') != body['description']:
            retry(service.events().update, calendarId=CAL_ID, eventId=g_evt_base['id'], body=body).execute()
            updated += 1
        else:
            skipped += 1
    else:
        # Single (non-recurring) event
        comp = base_comp
        s = comp['DTSTART'].dt
        e = comp.get('DTEND', comp['DTSTART']).dt
        ad = maybe_all_day(s, e)
        if ad:
            g_start = {'date': ad[0].isoformat()}
            g_end   = {'date': ad[1].isoformat()}
        else:
            if isinstance(s, datetime.date) and not isinstance(s, datetime.datetime):
                e = e if isinstance(e, datetime.date) else s + datetime.timedelta(days=1)
            g_start, g_end = to_gcal_dt(s), to_gcal_dt(e)
        body = {
            'summary':  str(comp.get('SUMMARY', 'No title')),
            'location': str(comp.get('LOCATION', '')),
            'start':    g_start,
            'end':      g_end,
            'extendedProperties': {'private': {'compositeUID': base_comp_key}}
        }
        body['description'] = body_hash(body)
        g_evt = google.get(base_comp_key)
        if not g_evt:
            retry(service.events().insert, calendarId=CAL_ID, body=body).execute()
            added += 1
        elif g_evt.get('description') != body['description']:
            retry(service.events().update, calendarId=CAL_ID, eventId=g_evt['id'], body=body).execute()
            updated += 1
        else:
            skipped += 1

# ───────────────────── delete missing / quarantined ──────────────────────
queue = deque([(u, g['id']) for u, g in google.items() if (u not in outlook) or (u.split('❄')[0] in quarantined)])
deleted = failed = 0
while queue:
    for _ in range(min(DELETE_BATCH, len(queue))):
        uid, eid = queue.popleft()
        try:
            retry(service.events().delete, calendarId=CAL_ID, eventId=eid).execute()
            deleted += 1
        except HttpError as e:
            if e.resp.status in (404, 410):
                deleted += 1
            elif e.resp.status in (403, 429):
                queue.append((uid, eid))
            else:
                failed += 1
    if queue:
        time.sleep(2)

# ───────────────────── summary ───────────────────────────────────────────
print(f"\nSync complete: ➕{added} 🔄{updated} ⏭{skipped} ❌{deleted} ⚠️failed:{failed}")
