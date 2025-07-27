#!/usr/bin/env python3
"""
Strip problem headers (e.g., X-ENTOURAGE_UUID) from every *.ics in outbox/.
Run **after** exportEvents.scpt, before safe_sync.py.
"""
import re, pathlib, sys

OUTBOX = pathlib.Path.home() / "calendarBridge" / "outbox"
BAD_HDR  = re.compile(r"^X-(ENTOURAGE|MS-OLK).*", re.I)

count = 0
for ics in OUTBOX.glob("*.ics"):
    text = ics.read_text(errors="ignore").splitlines()
    cleaned = [ln for ln in text if not BAD_HDR.match(ln)]
    ics.write_text("\n".join(cleaned))
    count += 1

print(f"🧽 Cleaned {count} ICS files")
sys.exit(0)
