#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from pprint import pprint

# Import our existing logic
from safe_sync import (
    load_local,
    to_body,
    TIMEZONE,
    CAL_ID,
    ROOT,
    service,
    find_google_by_icaluid,
)

# List of UIDs that blew up in your log.
# You can add/remove from here as needed.
TARGET_UIDS = [
    "040000008200E00074C5B7101A82E00800000000F2E0132DD47BDB01000000000000000010000000C3E9E1199638D243867C988CE4B1E27DMxZgBGAAAAAADtFzEWmEfyTJby1KKFPQbcBwA/X+xYuu/hTYKp/+t0gFB8AAAABMWNAAC0i2g44wveRYn7w+GHKTCsAAby8JrMAAA=",
    "040000008200E00074C5B7101A82E00800000000F1502C4F695CD001000000000000000010000000FF5D02447F4F5A48BEAFB7A60E7D1875MxZgFRAAgI0wy+VRMAAEYAAAAA7RcxFphH8kyW8tSihT0G3AcAP1/sWLrv4U2Cqf/rdIBQfAAAAATFjQAAfix1MNrK0kq/YbFK8dUGOACPgKoHrwAAEA==",
    "040000008200E00074C5B7101A82E0080000000010BC2BC8644ADB01000000000000000010000000CF5AC2076F7B6A46A8619C5531393ED3MxZgBGAAAAAADtFzEWmEfyTJby1KKFPQbcBwA/X+xYuu/hTYKp/+t0gFB8AAAABMWNAAC0i2g44wveRYn7w+GHKTCsAAbBuDJAAAA=",
]

def main():
    outbox = f"{ROOT}/outbox"
    local = load_local(outbox)

    for target_uid in TARGET_UIDS:
        print("=" * 80)
        print(f"UID: {target_uid}")

        # Find local event(s) with this UID
        matching = [
            (k, ev)
            for (k, ev) in local.items()
            if ev.get("uid") == target_uid
        ]

        if not matching:
            print("  [LOCAL] No local events found with this UID.")
            continue

        print(f"  [LOCAL] Found {len(matching)} local instance(s).")
        for key, ev in matching:
            print(f"\n  Instance key: {key}")
            print(f"    allDay: {ev['allDay']}")
            print(f"    start:  {ev['start']!r}")
            print(f"    end:    {ev['end']!r}")

            desired_body = to_body(ev, TIMEZONE)
            print("  [LOCAL] Desired Google body:")
            pprint(desired_body)

            # Now look up in Google by iCalUID
            g = find_google_by_icaluid(CAL_ID, target_uid)
            if not g:
                print("  [GOOGLE] No event found with this iCalUID.")
                continue

            print("  [GOOGLE] Existing event summary / id:")
            print(f"    summary: {g.get('summary')!r}")
            print(f"    id:      {g.get('id')!r}")
            print("  [GOOGLE] Existing start / end:")
            pprint(g.get("start"))
            pprint(g.get("end"))

if __name__ == "__main__":
    main()
