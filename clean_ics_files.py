#!/usr/bin/env python3
import os
import re

ROOT = os.path.expanduser("~/calendarBridge")
OUTBOX = os.path.join(ROOT, "outbox")
SOURCE = os.path.join(OUTBOX, "outlook_full_export.ics")

def clean_and_split():
    if not os.path.exists(SOURCE):
        print(f"[WARN] No source ICS found at {SOURCE}")
        return

    with open(SOURCE, "r", errors="ignore") as f:
        text = f.read()

    # Split by VCALENDAR boundaries
    blocks = re.split(r"(?=BEGIN:VCALENDAR)", text)
    blocks = [b.strip() for b in blocks if b.strip()]

    cleaned_files = []
    for i, block in enumerate(blocks, 1):
        # Remove non-standard headers Outlook injects
        block = re.sub(r"^X-.*\r?\n?", "", block, flags=re.MULTILINE)

        # Ensure it ends properly
        if not block.endswith("END:VCALENDAR"):
            block += "\nEND:VCALENDAR"

        target = os.path.join(OUTBOX, f"clean_{i}.ics")
        with open(target, "w") as out:
            out.write(block)
        cleaned_files.append(target)

    print(f"[INFO] Cleaned and split into {len(cleaned_files)} ICS files")
    for f in cleaned_files:
        print(f"  - {f}")

if __name__ == "__main__":
    clean_and_split()
