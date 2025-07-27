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
            if hasattr(e, "resp") and getattr(e.resp, "status", 0) in (403, 429):
                time.sleep(5)
            else:
                raise

# ───────────────────── Google auth / build ───────────────────────────────
from google.oauth2.credentials  import Credentials
from google_auth_oauthlib.flow  import InstalledAppFlow
from googleapiclient.discovery  import build
from googleapiclient.errors     import HttpError
from google.auth.transport.requests import Request

BASE   = os.path.expanduser("~/calendarBridge")
TOKEN  = f"{BASE}/token.json"
CREDS  = f"{BASE}/credentials.json"
SCOPES = ["https://www.googleapis.com/auth/calendar"]
CAL_ID = "primary"

creds = None
if os.path.exists(TOKEN):
    creds = Credentials.from_authorized_user_file(TOKEN, SCOPES)
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        creds = InstalledAppFlow.from_client_secrets_file(CREDS, SCOPES) \
                                .run_local_server(port=0)
    with open(TOKEN, "w") as f:
        f.write(creds.to_json())

service = build("calendar", "v3", credentials=creds, cache_discovery=False)

# ───────────────────── read Outlook .ics files ───────────────────────────
from icalendar import Calendar
ICS_DIR = f"{BASE}/outbox"
outlook = {}
for fn in os.listdir(ICS_DIR):
    if fn.endswith(".ics"):
        with open(f"{ICS_DIR}/{fn}", "rb") as f:
            cal = Calendar.from_ical(f.read())
            for ev in cal.walk("VEVENT"):
                uid = str(ev.get("UID"))
                rec = str(ev.get("RECURRENCE-ID", ""))
                outlook[f"{uid}❄{rec}"] = ev
print(f"📂 Parsed {len(outlook)} Outlook events")

# ───────────────────── pull Google events we manage ──────────────────────
google, page = {}, None
while True:
    resp = retry(service.events().list,
                 calendarId=CAL_ID,
                 pageToken=page,
                 maxResults=2500,
                 fields=("nextPageToken,"
                         "items(id,summary,description,start,end,extendedProperties)")).execute()
    for it in resp.get("items", []):
        key = it.get("extendedProperties", {}).get("private", {}).get("compositeUID")
        if key:
            google[key] = it
    page = resp.get("nextPageToken")
    if not page:
        break
print(f"☁️  Loaded {len(google)} Google events tagged as Outlook-origin")

# ───────────────────── helpers ───────────────────────────────────────────
def to_gcal_dt(dt):
    return {"dateTime": dt.astimezone().isoformat()} if isinstance(dt, datetime.datetime) else {"date": dt.isoformat()}

def body_hash(b):
    return hashlib.sha256(json.dumps(b, sort_keys=True).encode()).hexdigest()

def maybe_all_day(start, end):
    """If span is midnight-to-midnight whole-days, return (d1,d2) else None."""
    if isinstance(start, datetime.datetime) and isinstance(end, datetime.datetime):
        if start.time() == end.time() == datetime.time(0, 0) and (end - start).seconds == 0:
            return start.date(), end.date()
    return None

DELETE_BATCH = 50
QUAR_FILE    = f"{BASE}/quarantine.txt"
quarantined  = set(open(QUAR_FILE).read().split()) if os.path.exists(QUAR_FILE) else set()

# ───────────────────── add / update phase ────────────────────────────────
added = updated = skipped = 0
for uid, comp in outlook.items():
    if uid in quarantined:
        continue

    s = comp["DTSTART"].dt
    e = comp.get("DTEND", comp["DTSTART"]).dt

    # figure Google start / end
    ad = maybe_all_day(s, e)
    if ad:
        g_start = {"date": ad[0].isoformat()}
        g_end   = {"date": ad[1].isoformat()}
    else:
        if isinstance(s, datetime.date) and not isinstance(s, datetime.datetime):
            e = e if isinstance(e, datetime.date) else s + datetime.timedelta(days=1)
        g_start, g_end = to_gcal_dt(s), to_gcal_dt(e)

    body = {
        "summary":  str(comp.get("SUMMARY", "No title")),
        "location": str(comp.get("LOCATION", "")),
        "start":    g_start,
        "end":      g_end,
        "extendedProperties": {"private": {"compositeUID": uid}}
    }
    body["description"] = body_hash(body)

    g_evt = google.get(uid)
    if not g_evt:
        retry(service.events().insert, calendarId=CAL_ID, body=body).execute()
        added += 1
    elif g_evt.get("description") != body["description"]:
        retry(service.events().update, calendarId=CAL_ID,
              eventId=g_evt["id"], body=body).execute()
        updated += 1
    else:
        skipped += 1

# ───────────────────── delete missing / quarantined ──────────────────────
queue = deque([(u, g["id"]) for u, g in google.items()
               if (u not in outlook) or (u in quarantined)])
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
