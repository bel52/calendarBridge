#!/usr/bin/env python3
"""
safe_sync.py  –  Outlook ➜ Google Calendar one-way sync
• pulls Outlook .ics from ~/calendarBridge/outbox
• pushes/updates/deletes on Google Calendar ‘primary’
• keeps a 60 s socket timeout & automatic retries
• NO longer uses the broken privateExtendedProperty filter
"""

import os, json, time, socket, hashlib, datetime
from collections import deque

# ───── global socket timeout ─────────────────────────────────────────────
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
            if getattr(e, "resp", None) and getattr(e.resp, "status", 0) in (403, 429):
                time.sleep(5)
            else:
                raise

def gexec(req):
    return retry(lambda: req.execute())

# ───── Google auth ───────────────────────────────────────────────────────
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery  import build
from googleapiclient.errors     import HttpError
from google.auth.transport.requests import Request      # refresh helper

BASE   = os.path.expanduser("~/calendarBridge")
TOKEN  = f"{BASE}/token.json"
CREDS  = f"{BASE}/credentials.json"
SCOPES = ["https://www.googleapis.com/auth/calendar"]
CAL_ID = "primary"

creds = Credentials.from_authorized_user_file(TOKEN, SCOPES) if os.path.exists(TOKEN) else None
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        creds = InstalledAppFlow.from_client_secrets_file(CREDS, SCOPES)\
                                .run_local_server(port=0)
    with open(TOKEN, "w") as fh:
        fh.write(creds.to_json())

service = build("calendar", "v3", credentials=creds, cache_discovery=False)

# ───── read Outlook .ics ─────────────────────────────────────────────────
from icalendar import Calendar
ICS_DIR = f"{BASE}/outbox"
outlook = {}
for fn in os.listdir(ICS_DIR):
    if fn.endswith(".ics"):
        with open(f"{ICS_DIR}/{fn}", "rb") as fh:
            cal = Calendar.from_ical(fh.read())
            for ev in cal.walk("VEVENT"):
                uid  = str(ev.get("UID"))
                rec  = str(ev.get("RECURRENCE-ID", ""))
                outlook[f"{uid}❄{rec}"] = ev
print(f"📂 Parsed {len(outlook)} Outlook events")

# ───── fetch Google events (no extended-property filter) ─────────────────
google = {}
page = None
while True:
    resp = gexec(service.events().list(
        calendarId=CAL_ID,
        pageToken=page,
        maxResults=2500,
        fields="nextPageToken,items(id,summary,description,start,end,extendedProperties)"
    ))
    for item in resp.get("items", []):
        priv = item.get("extendedProperties", {}).get("private", {})
        key  = priv.get("compositeUID")
        if key:
            google[key] = item
    page = resp.get("nextPageToken")
    if not page:
        break
print(f"☁️  Loaded {len(google)} Google events tagged as Outlook-origin")

# ───── helpers ───────────────────────────────────────────────────────────
def to_gcal_dt(dt):
    return {"dateTime": dt.astimezone().isoformat()} if isinstance(dt, datetime.datetime) \
           else {"date": dt.isoformat()}

def body_hash(b):
    return hashlib.sha256(json.dumps(b, sort_keys=True).encode()).hexdigest()

DELETE_BATCH   = 50
QUAR_FILE      = f"{BASE}/quarantine.txt"
quarantined    = set(open(QUAR_FILE).read().split()) if os.path.exists(QUAR_FILE) else set()

# ───── add / update ──────────────────────────────────────────────────────
added = updated = skipped = 0
for uid, comp in outlook.items():
    if uid in quarantined:
        continue
    s = comp["DTSTART"].dt
    e = comp.get("DTEND", comp["DTSTART"]).dt
    if isinstance(s, datetime.date) and not isinstance(s, datetime.datetime):
        e = e if isinstance(e, datetime.date) else s + datetime.timedelta(days=1)

    body = {
        "summary":  str(comp.get("SUMMARY", "No title")),
        "location": str(comp.get("LOCATION", "")),
        "start":    to_gcal_dt(s),
        "end":      to_gcal_dt(e),
        "extendedProperties": {"private": {"compositeUID": uid}}
    }
    body["description"] = body_hash(body)

    g_evt = google.get(uid)
    if not g_evt:
        gexec(service.events().insert(calendarId=CAL_ID, body=body))
        added += 1
    elif g_evt.get("description") != body["description"]:
        gexec(service.events().update(calendarId=CAL_ID,
                                      eventId=g_evt["id"], body=body))
        updated += 1
    else:
        skipped += 1

# ───── delete missing / quarantined ──────────────────────────────────────
from collections import deque
queue = deque([
    (uid, g["id"]) for uid, g in google.items()
    if (uid not in outlook) or (uid in quarantined)
])
deleted = failed = 0
while queue:
    for _ in range(min(DELETE_BATCH, len(queue))):
        uid, eid = queue.popleft()
        try:
            gexec(service.events().delete(calendarId=CAL_ID, eventId=eid))
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

# ───── summary ───────────────────────────────────────────────────────────
print(f"\nSync complete: ➕{added} 🔄{updated} ⏭{skipped} ❌{deleted} ⚠️failed:{failed}")
