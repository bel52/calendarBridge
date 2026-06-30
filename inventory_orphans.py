"""
inventory_orphans.py - READ-ONLY.
Lists Google events in the sync window with no matching Outlook event,
and reports whether CalendarBridge would consider each deletable.
Performs NO writes.
"""
import safe_sync as sb


def main():
    cfg = sb.load_config()
    tz = sb.ZoneInfo(cfg["timezone"])
    cal_id = cfg.get("calendar_id") or cfg.get("calendarId") or "primary"

    service = sb.get_google_service()
    google_events = sb.fetch_google_events(service, cal_id, cfg, tz)
    local_events = sb.load_local_events(cfg, tz)

    local_keys = set(local_events.keys())
    orphans = [(k, it) for k, it in google_events.items() if k not in local_keys]

    print()
    print(f"Local (Outlook) events:                {len(local_events)}")
    print(f"Google events (with usable iCalUID):   {len(google_events)}")
    print(f"Google-only (orphan candidates):       {len(orphans)}")
    print("=" * 100)

    def start_of(it):
        s = it.get("start") or {}
        return s.get("dateTime") or s.get("date") or ""

    tagged = untagged = 0
    for key, item in sorted(orphans, key=lambda x: start_of(x[1])):
        ext = (item.get("extendedProperties") or {}).get("private") or {}
        is_ours = ext.get("source") == sb.ORPHAN_MARKER
        if is_ours:
            tagged += 1
        else:
            untagged += 1
        start = start_of(item) or "?"
        summary = (item.get("summary") or "(no title)")[:45]
        gid = item.get("id", "?")
        flag = "DELETABLE" if is_ours else "SKIPPED (no marker)"
        print(f"{start:25}  {summary:47}  {flag}")
        print(f"{'':25}  gid={gid}")

    print("=" * 100)
    print(f"Deletable by sync (source=CalendarBridge): {tagged}")
    print(f"Skipped by sync (untagged strays):         {untagged}")
    print("No changes were made. This was read-only.")
    print()


if __name__ == "__main__":
    main()
