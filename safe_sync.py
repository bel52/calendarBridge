#!/usr/bin/env python3
"""
safe_sync.py – Outlook ➜ Google Calendar one-way sync
- Explicit timeZone on dateTime events
- RRULE sanitizer for Google
- EXDATEs for modified & cancelled recurring instances (prevents duplicates)
- Python 3.9 compatible
"""
import os, json, time, socket, hashlib, datetime, re
from collections import deque, defaultdict
from typing import Optional
from zoneinfo import ZoneInfo

socket.setdefaulttimeout(60)

def retry(fn, *a, **kw):
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

# ── Config ────────────────────────────────────────────────────────────────
BASE   = os.path.expanduser('~/calendarBridge')
TOKEN  = f"{BASE}/token.json"
CREDS  = f"{BASE}/credentials.json"
SCOPES = ['https://www.googleapis.com/auth/calendar']
CAL_ID = 'primary'
DEFAULT_TZ_NAME = os.environ.get('CALBRIDGE_TZ', 'America/New_York')
try:
    DEFAULT_TZ = ZoneInfo(DEFAULT_TZ_NAME)
except Exception:
    DEFAULT_TZ_NAME = 'UTC'
    DEFAULT_TZ = ZoneInfo('UTC')

# ── Google auth ───────────────────────────────────────────────────────────
from google.oauth2.credentials  import Credentials
from google_auth_oauthlib.flow  import InstalledAppFlow
from googleapiclient.discovery  import build
from googleapiclient.errors     import HttpError
from google.auth.transport.requests import Request

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

# ── Read Outlook .ics files ───────────────────────────────────────────────
from icalendar import Calendar
ICS_DIR = f"{BASE}/outbox"
outlook = {}
all_components = []  # keep raw list for cancellation tracking
for fn in os.listdir(ICS_DIR):
    if fn.endswith('.ics'):
        with open(f"{ICS_DIR}/{fn}", 'rb') as f:
            cal = Calendar.from_ical(f.read())
            for ev in cal.walk('VEVENT'):
                all_components.append(ev)
                uid = str(ev.get('UID'))
                rec = str(ev.get('RECURRENCE-ID', ''))
                outlook[f"{uid}❄{rec}"] = ev
print(f"📂 Parsed {len(outlook)} Outlook events")

# Track cancelled instances by base UID so we can add EXDATE (but don't create events)
cancelled_by_base = defaultdict(list)
for comp in list(outlook.values()):
    if str(comp.get('STATUS', '')).upper() == 'CANCELLED':
        uid = str(comp.get('UID'))
        rec_prop = comp.get('RECURRENCE-ID')
        if rec_prop:
            cancelled_by_base[uid].append(rec_prop)
        # Remove so we won't try to create an event for it
        outlook.pop(f"{uid}❄{str(rec_prop)}", None)

# ── Load Google events we manage ──────────────────────────────────────────
google, page = {}, None
while True:
    resp = retry(service.events().list,
                 calendarId=CAL_ID,
                 pageToken=page,
                 maxResults=2500,
                 fields='nextPageToken,items(id,summary,description,start,end,extendedProperties)').execute()
    for it in resp.get('items', []):
        key = it.get('extendedProperties', {}).get('private', {}).get('compositeUID')
        if key:
            google[key] = it
    page = resp.get('nextPageToken')
    if not page:
        break
print(f"☁️  Loaded {len(google)} Google events tagged as Outlook-origin")

# ── Helpers ───────────────────────────────────────────────────────────────
def ensure_tz(dt: datetime.datetime) -> datetime.datetime:
    if isinstance(dt, datetime.datetime) and (dt.tzinfo is None or dt.tzinfo.utcoffset(dt) is None):
        return dt.replace(tzinfo=DEFAULT_TZ)
    return dt

def to_gcal_dt(dt):
    if isinstance(dt, datetime.datetime):
        dt = ensure_tz(dt)
        return {'dateTime': dt.isoformat(), 'timeZone': DEFAULT_TZ_NAME}
    else:
        return {'date': dt.isoformat()}

def body_hash(b):
    return hashlib.sha256(json.dumps(b, sort_keys=True).encode()).hexdigest()

def maybe_all_day(start, end):
    if isinstance(start, datetime.datetime) and isinstance(end, datetime.datetime):
        start = ensure_tz(start); end = ensure_tz(end)
        if start.time() == end.time() == datetime.time(0, 0) and (end - start).seconds == 0:
            return start.date(), end.date()
    return None

# Minimal RRULE sanitizer for Google
_ALLOWED_KEYS = {"FREQ","INTERVAL","BYDAY","BYMONTHDAY","BYMONTH","COUNT","UNTIL","WKST","BYSETPOS","BYHOUR","BYMINUTE","BYSECOND","BYWEEKNO","BYYEARDAY"}
_DT_RE   = re.compile(r'^\d{8}T\d{6}$')
_DT_Z_RE = re.compile(r'^\d{8}T\d{6}Z$')
_DATE_RE = re.compile(r'^\d{8}$')

def sanitize_rrule(rrule_prop) -> Optional[str]:
    if not rrule_prop:
        return None
    r = rrule_prop.to_ical().decode() if hasattr(rrule_prop, 'to_ical') else str(rrule_prop)
    r = r.replace('\r','').replace('\n','').strip()
    if r.upper().startswith('RRULE:'):
        r = r[6:]
    parts = [p for p in r.split(';') if p and '=' in p]
    out = []
    for p in parts:
        k, v = p.split('=', 1)
        kU = k.strip().upper()
        vU = v.strip().upper()
        if kU not in _ALLOWED_KEYS:
            continue
        if kU == "COUNT":
            if not vU.isdigit() or int(vU) < 1:
                continue
        if kU == "INTERVAL":
            try:
                if int(vU) < 1:
                    continue
            except:
                continue
        if kU == "UNTIL":
            if _DATE_RE.fullmatch(vU):
                pass
            elif _DT_Z_RE.fullmatch(vU):
                pass
            elif _DT_RE.fullmatch(vU):
                vU = vU + "Z"
            else:
                continue
        if kU == "BYDAY" and vU == "":
            continue
        out.append(f"{kU}={vU}")
    if not any(x.startswith("FREQ=") for x in out):
        return None
    return "RRULE:" + ";".join(out)

def exdate_line_from_dt(dt) -> str:
    """Build an EXDATE line in UTC (Z) or DATE for all-day."""
    if isinstance(dt, datetime.datetime):
        dt = ensure_tz(dt).astimezone(datetime.timezone.utc)
        return "EXDATE:" + dt.strftime("%Y%m%dT%H%M%SZ")
    else:
        # all-day date
        return "EXDATE:" + dt.strftime("%Y%m%d")

DELETE_BATCH = 50
QUAR_FILE    = f"{BASE}/quarantine.txt"
quarantined  = set(open(QUAR_FILE).read().split()) if os.path.exists(QUAR_FILE) else set()

# ── Add / Update ──────────────────────────────────────────────────────────
added = updated = skipped = 0

# Group by base UID to handle recurring with exceptions
grouped = {}
for comp_key, comp in outlook.items():
    base_uid = comp_key.split('❄')[0]
    grouped.setdefault(base_uid, []).append(comp_key)

for base_uid, comp_keys in grouped.items():
    base_comp_key = base_uid + '❄'
    base_comp = outlook.get(base_comp_key)
    if base_comp is None or base_uid in quarantined:
        continue

    if base_comp.get('RRULE'):
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

        # Recurrence: RRULE + EXDATEs for modified & cancelled instances
        recurrences = []
        rrule_str = sanitize_rrule(base_comp.get('RRULE'))
        if rrule_str:
            recurrences.append(rrule_str)

        # EXDATE for CANCELLED instances
        for rec_prop in cancelled_by_base.get(base_uid, []):
            try:
                dt = rec_prop.dt
                recurrences.append(exdate_line_from_dt(dt))
            except Exception:
                pass

        # Handle modified instances: add EXDATE for each, and create separate event
        for comp_key in comp_keys:
            if comp_key == base_comp_key:
                continue
            exc_comp = outlook[comp_key]
            rec_prop = exc_comp.get('RECURRENCE-ID')
            if rec_prop:
                try:
                    recurrences.append(exdate_line_from_dt(rec_prop.dt))
                except Exception:
                    pass

            status = str(exc_comp.get('STATUS', ''))
            if status.upper() == 'CANCELLED':
                continue  # already excluded via EXDATE above

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

# ── Delete missing / quarantined ──────────────────────────────────────────
queue = deque([(u, g['id']) for u, g in google.items() if (u not in outlook) or (u.split('❄')[0] in quarantined)])
deleted = failed = 0
while queue:
    for _ in range(min(50, len(queue))):
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

print(f\"\\nSync complete: ➕{added} 🔄{updated} ⏭{skipped} ❌{deleted} ⚠️failed:{failed}\")
