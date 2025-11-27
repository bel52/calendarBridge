import datetime
import os

from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# Same scope you're already using for safe_sync.py
SCOPES = ["https://www.googleapis.com/auth/calendar"]


def get_service():
    token_path = os.path.join(os.path.dirname(__file__), "token.json")
    creds = Credentials.from_authorized_user_file(token_path, SCOPES)
    return build("calendar", "v3", credentials=creds)


def main():
    service = get_service()

    today = datetime.date.today()

    event = {
        "summary": "TEST all-day banner (delete me)",
        "description": "Inserted by test_all_day_event.py using date-only fields",
        # KEY PART: use `date`, NOT `dateTime`
        "start": {
            "date": today.isoformat()
        },
        "end": {
            # end is exclusive; this makes it a single-day all-day event
            "date": (today + datetime.timedelta(days=1)).isoformat()
        },
    }

    created = service.events().insert(calendarId="primary", body=event).execute()
    print("Created event:")
    print("  id:      ", created.get("id"))
    print("  htmlLink:", created.get("htmlLink"))


if __name__ == "__main__":
    main()
