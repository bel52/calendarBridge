#!/usr/bin/env python3
"""
safe_sync.py – Outlook → Google Calendar
• Same logic as before, plus rate-limit-safe deletion batching.
"""

import os, glob, hashlib, datetime, sys, time, pytz
from collections import deque
from icalendar import Calendar
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from googleapiclient.errors import HttpError

# ── CONFIG ───────────────────────────────────────────────────────────────
HOME        = os.path.expanduser("~")
BOX         = os.path.join(HOME, "calendarBridge", "outbox")
TOKEN_FILE  = os.path.join(HOME, "calendarBridge", "token.json")
CREDS_FILE  = os.path.join(HOME, "calendarBridge", "credentials.json")
CALENDAR_ID = "primary"
SCOPES      = ["https://www.googleapis.com/auth/calendar"]
TZ_FALLBACK = "America/New_York"
DELETE_BATCH = 80          # delete this many, then pause
# ─────────────────────────────────────────────────────────────────────────


def composite_uid(comp):
    uid = str(comp["UID"])
    rid = comp.get("RECURRENCE-ID")
    return f"{uid}|{rid.dt.isoformat() if rid else ''}"


def to_gcal_dt(dt):
    if isinstance(dt, datetime.datetime):
        if dt.tzinfo is None:
            dt = pytz.timezone(TZ_FALLBACK).localize(dt)
        return {"dateTime": dt.astimezone(pytz.utc).isoformat()}
    return {"date": dt.isoformat()}


# ── Authorise ────────────────────────────────────────────────────────────
creds = Credentials.from_authorized_user_file(TOKEN_FILE, SCOPES) \
        if os.path.exists(TOKEN_FILE) else None
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        creds = InstalledAppFlow.from_client_secrets_file(CREDS_FILE, SCOPES) \
                                .run_local_server(port=0)
    open(TOKEN_FILE, "w").write(creds.to_json())

service = build("calendar", "v3", credentials=creds)

# ── Parse Outlook *.ics* ─────────────────────────────────────────────────
outlook = {}
for p in glob.glob(os.path.join(BOX, "*.ics")):
    cal = Calendar.from_ical(open(p, "rb").read())
    for c in filter(lambda x: x.name == "VEVENT", cal.walk()):
        outlook[composite_uid(c)] = c
print(f"📂 Parsed {len(outlook)} Outlook events")

# ── Fetch Google events incl. legacy keys ────────────────────────────────
google = {}
page = None
while True:
    r = service.events().list(calendarId=CALENDAR_ID,
                              maxResults=2500, singleEvents=False,
                              pageToken=page,
                              fields=("nextPageToken,items(id,summary,"
                                      "extendedProperties,start,end,"
                                      "description)")).execute()
    for item in r.get("items", []):
        priv = item.get("extendedProperties", {}).get("private", {})
        key  = priv.get("compositeUID") or priv.get("icalUID") or priv.get("outlookUID")
        if key:
            google[key] = item
    page = r.get("nextPageToken")
    if not page:
        break
print(f"☁️  Loaded {len(google)} Google events tagged as Outlook-origin")

# ── Sync adds/updates (unchanged from prior version) ─────────────────────
added = updated = skipped = 0
for k, comp in outlook.items():
    s = comp["DTSTART"].dt
    e = comp.get("DTEND", comp["DTSTART"]).dt
    if isinstance(s, datetime.date) and not isinstance(s, datetime.datetime):
        if not isinstance(e, datetime.date):
            e = s + datetime.timedelta(days=1)
    body = {
        "summary":  str(comp.get("SUMMARY", "No title")),
        "location": str(comp.get("LOCATION", "")),
        "start":    to_gcal_dt(s),
        "end":      to_gcal_dt(e),
        "extendedProperties": {"private": {"compositeUID": k}}
    }
    h = hashlib.sha256(repr(body).encode()).hexdigest()
    body["description"] = h
    g = google.get(k)
    if not g:
        service.events().insert(calendarId=CALENDAR_ID, body=body).execute()
        added += 1
    else:
        if g.get("description") != h:
            service.events().update(calendarId=CALENDAR_ID,
                                    eventId=g["id"], body=body).execute()
            updated += 1
        else:
            skipped += 1

# ── Batched deletions with back-off ──────────────────────────────────────
to_delete = [ (k,g["id"],g.get("summary","")) for k,g in google.items()
              if k not in outlook ]
delete_queue = deque(to_delete)
deleted = failed = 0

while delete_queue:
    for _ in range(min(DELETE_BATCH, len(delete_queue))):
        k, eid, title = delete_queue.popleft()
        try:
            service.events().delete(calendarId=CALENDAR_ID, eventId=eid).execute()
            deleted += 1
        except HttpError as err:
            if err.resp.status in (404, 410):
                deleted += 1
            elif err.resp.status in (403, 429):
                # Push back for retry later
                delete_queue.append((k, eid, title))
            else:
                print(f"⚠️  Could not delete {title}: {err}")
                failed += 1
    if delete_queue:
        # Exponential back-off
        backoff = min(8, 2 ** (3 - (len(delete_queue) // DELETE_BATCH)))
        time.sleep(backoff)

# ── Summary ──────────────────────────────────────────────────────────────
print(f"\nSync complete: ➕{added} 🔄{updated} ⏭{skipped} ❌{deleted} ⚠️failed:{failed}")
sys.exit(0)
