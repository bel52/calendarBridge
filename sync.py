#!/usr/bin/env python3
import json, pathlib, textwrap
from icalendar import Calendar
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

BASE     = pathlib.Path.home() / "calendarBridge"
ICS_DIR  = BASE / "outbox"
CACHE    = BASE / "cache.json"
SCOPES   = ["https://www.googleapis.com/auth/calendar"]
CAL_ID   = "primary"                 # change to a secondary ID if you like

def gsvc():
    tok = BASE / "token.json"
    creds = Credentials.from_authorized_user_file(tok, SCOPES) if tok.exists() else None
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(BASE / "credentials.json", SCOPES)
        creds = flow.run_local_server(port=0)
        tok.write_text(creds.to_json())
    return build("calendar", "v3", credentials=creds)

def main():
    svc   = gsvc()
    cache = json.loads(CACHE.read_text()) if CACHE.exists() else {}

    inserted, updated = 0, 0
    for ics in sorted(ICS_DIR.glob("ev*.ics")):
        cal = Calendar.from_ical(ics.read_bytes())
        evt = cal.walk("vevent")[0]
        uid = str(evt["uid"])

        body = {
            "summary": str(evt.get("summary", "")),
            "start": {"dateTime": evt.decoded("dtstart").isoformat(),
                      "timeZone": "America/New_York"},
            "end":   {"dateTime": evt.decoded("dtend").isoformat(),
                      "timeZone": "America/New_York"},
        }

        if uid in cache:                              # update
            svc.events().update(calendarId=CAL_ID,
                                 eventId=cache[uid], body=body).execute()
            updated += 1
        else:                                         # first insert
            ev = svc.events().insert(calendarId=CAL_ID, body=body).execute()
            cache[uid] = ev["id"]
            inserted += 1

    CACHE.write_text(json.dumps(cache))
    print(f"✔︎ inserted {inserted}, updated {updated}")

if __name__ == "__main__":
    main()
