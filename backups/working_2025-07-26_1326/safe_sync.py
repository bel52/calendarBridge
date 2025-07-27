#!/usr/bin/env python3
"""
safe_sync.py – Outlook → Google Calendar one-way bridge

• Keys are just VEVENT UID (stable)
• Global 120-s socket timeout + 3-try retry wrapper
• Skips any UID listed in quarantine.txt
"""

import os, json, time, socket, hashlib, datetime
from collections import deque

socket.setdefaulttimeout(120)

def retry(fn):
    for n in range(3):
        try:
            return fn()
        except (socket.timeout, TimeoutError):
            if n == 2: raise
            time.sleep(5)
        except Exception as e:
            if getattr(getattr(e, "resp", None), "status", 0) in (403, 429):
                time.sleep(5)
            else:
                raise

gexec = lambda req: retry(req.execute)

# ───────────── Google auth ───────────────────────────────────────────────
from google.oauth2.credentials      import Credentials
from google_auth_oauthlib.flow      import InstalledAppFlow
from googleapiclient.discovery      import build
from googleapiclient.errors         import HttpError
from google.auth.transport.requests import Request

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
    open(TOKEN, "w").write(creds.to_json())

service = build("calendar", "v3", credentials=creds, cache_discovery=False)

# ───────────── Load Outlook .ics (keep-first per UID) ────────────────────
from icalendar import Calendar
ICS_DIR = f"{BASE}/outbox"
outlook = {}

for fn in os.listdir(ICS_DIR):
    if not fn.endswith(".ics"): continue
    with open(f"{ICS_DIR}/{fn}", "rb") as f:
        cal = Calendar.from_ical(f.read())
        for ev in cal.walk("VEVENT"):
            uid = str(ev.get("UID"))
            if uid not in outlook:           # keep first occurrence
                outlook[uid] = ev
print(f"📂 Parsed {len(outlook)} Outlook events")

# ───────────── Pull Google events tagged with compositeUID ───────────────
google = {}
page = None
while True:
    resp = gexec(service.events().list(
        calendarId=CAL_ID,
        pageToken=page,
        maxResults=2500,
        fields=("nextPageToken,"
                "items(id,summary,description,start,end,extendedProperties)")))
    for it in resp.get("items", []):
        key = it.get("extendedProperties", {}).get("private", {}).get("compositeUID")
        if key:
            google[key] = it
    page = resp.get("nextPageToken")
    if not page:
        break
print(f"☁️  Loaded {len(google)} Google events tagged as Outlook-origin")

# ───────────── helpers ───────────────────────────────────────────────────
def to_gcal_dt(dt):
    return {"dateTime": dt.astimezone().isoformat()} \
           if isinstance(dt, datetime.datetime) else {"date": dt.isoformat()}

body_hash = lambda b: hashlib.sha256(json.dumps(b, sort_keys=True).encode()).hexdigest()

QUAR_FILE   = f"{BASE}/quarantine.txt"
quarantined = set(open(QUAR_FILE).read().split()) if os.path.exists(QUAR_FILE) else set()

DELETE_BATCH = 50

# ───────────── Add / Update ──────────────────────────────────────────────
added = updated = skipped = 0
for uid, comp in outlook.items():
    if uid in quarantined: continue

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
        gexec(service.events().update(calendarId=CAL_ID, eventId=g_evt["id"], body=body))
        updated += 1
    else:
        skipped += 1

# ───────────── Delete missing / quarantined ──────────────────────────────
queue = deque([(uid, g["id"]) for uid, g in google.items()
               if (uid not in outlook) or (uid in quarantined)])

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

# ───────────── Summary ───────────────────────────────────────────────────
print(f"\nSync complete: ➕{added} 🔄{updated} ⏭{skipped} ❌{deleted} ⚠️failed:{failed}")
