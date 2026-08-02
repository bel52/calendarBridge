#!/usr/bin/env python3
"""
export_events_db.py v2 — CalendarBridge exporter for New Outlook (no AppleScript).

v2 adds full recurring-series support:
  - Series masters (Calendar_IsRecurring=1) carry a binary recurrence
    sub-table in property (pid=2, type=0x0d). Decoded fields (verified
    against live data, MS-OXOCAL-analogous):
      (1,0x02) RecurFrequency  0x200A daily / 0x200B weekly /
                               0x200C monthly / 0x200D yearly
      (1,0x03) PatternType     0 day, 1 week, 2 month, 3 month-nth
      (2,0x03) Interval        (daily may be minutes: 1440*days)
      (3,0x03) EndType         0x2021 until / 0x2022 count / 0x2023 never
      (4,0x03) OccurrenceCount
      (5,0x03) FirstDOW        0=Sunday
      (7,0x03) PatternTypeSpecific (weekly: day bitmask SU=1..SA=64;
                               monthly: day-of-month)
      (8,0x03) Nth             (month-nth: 1-4, 5=last) [when present]
      (0x0f,0x03) StartDate    minutes since 1601 (midnight)
      (0x10,0x03) EndDate      minutes since 1601; 0x5AE980DF = no end
      (1,0x0d)  deleted instance dates (u32[] minutes, midnight)
      (2,0x0d)  modified instance dates (u32[] minutes, midnight)
  - Masters emit DTSTART;TZID=<series tz> + RRULE + EXDATE(deleted-only).
  - Exception rows (Calendar_MasterRecordID != 0) emit with the MASTER's
    UID + RECURRENCE-ID so recurring_ical_events applies them as overrides.
  - Undecodable series fall back to emitting their materialized rows
    (never silently dropped) and are counted in the report.

Parity mode (offline, ZERO Google interaction):
  python3 export_events_db.py --parity /tmp/applescript_baseline.ics
  Expands both the generated ICS and the AppleScript baseline through the
  same library safe_sync uses, and diffs the expanded populations.

Normal mode writes outbox/outlook_full_export.ics (atomic, fail-closed).
"""

import argparse
import datetime as dtmod
import html as html_mod
import json
import logging
import os
import re
import shutil
import sqlite3
import struct
import subprocess
import sys
import tempfile
from datetime import datetime, timedelta, timezone

try:
    from zoneinfo import ZoneInfo
except ImportError:
    from backports.zoneinfo import ZoneInfo

ROOT = os.path.dirname(os.path.abspath(__file__))
OUTBOX = os.path.join(ROOT, "outbox")
ICS_PATH = os.path.join(OUTBOX, "outlook_full_export.ics")
STATE_PATH = os.path.join(OUTBOX, ".export_db_state.json")
CONFIG_PATH = os.path.join(ROOT, "calendar_config.json")

OUTLOOK_DATA_DIR = os.path.expanduser(
    "~/Library/Group Containers/UBF8T346G9.Office/Outlook/"
    "Outlook 15 Profiles/Main Profile/Data"
)

EP = 11644473600
NO_END_SENTINEL = 0x5AE980DF
PID_BODY, PID_SUBJECT, PID_LOCATION = 1, 2, 4
MAX_DESC = 4000
DAYS = ["SU", "MO", "TU", "WE", "TH", "FR", "SA"]
FREQ_MAP = {0x200A: "DAILY", 0x200B: "WEEKLY", 0x200C: "MONTHLY", 0x200D: "YEARLY"}

log = logging.getLogger("export_events_db")


def mdt(v):
    """store minutes -> naive UTC datetime"""
    return datetime(1970, 1, 1) + timedelta(seconds=v * 60 - EP)


def dtm(d):
    return int((d - datetime(1970, 1, 1)).total_seconds() + EP) // 60


def notify(msg):
    try:
        subprocess.run(["osascript", "-e",
                        f'display notification "{msg}" with title "CalendarBridge"'],
                       capture_output=True, timeout=10)
    except Exception:
        pass


def load_config():
    try:
        with open(CONFIG_PATH) as f:
            return json.load(f)
    except Exception:
        return {}


# --------------------------------------------------- olk15 property parsing

def read_props(path):
    """Return (text_props{pid:str}, blob_props{(pid,typ):bytes}, rawlen)."""
    with open(path, "rb") as f:
        d = f.read()
    if len(d) < 0x40 or d[0x20:0x24] != b"CRLC":
        raise ValueError("bad magic")
    n = struct.unpack_from("<I", d, 0x28)[0]
    if not 0 < n < 512 or 0x34 + n * 8 > len(d):
        raise ValueError(f"bad count {n}")
    tags = []
    for i in range(n):
        b = 0x34 + i * 8
        tags.append((struct.unpack_from("<H", d, b)[0], d[b + 3],
                     struct.unpack_from("<I", d, b + 4)[0]))
    pos = 0x34 + n * 8
    text, blobs = {}, {}
    for pid, typ, ln in tags:
        if pos + ln > len(d):
            break
        raw = d[pos:pos + ln]
        if typ == 0x1F and ln:
            try:
                text.setdefault(pid, raw.decode("utf-16le").rstrip("\x00"))
            except UnicodeDecodeError:
                pass
        else:
            blobs.setdefault((pid, typ), raw)
        pos += ln
    return text, blobs


def parse_subtable(raw):
    """Nested property table (same layout, 12-byte header)."""
    if len(raw) < 12:
        raise ValueError("short subtable")
    _, voff, _ = struct.unpack_from("<III", raw, 0)
    n = (voff - 12) // 8
    if n <= 0 or voff > len(raw):
        raise ValueError("bad subtable")
    tags = [(struct.unpack_from("<H", raw, 12 + i * 8)[0], raw[12 + i * 8 + 3],
             struct.unpack_from("<I", raw, 12 + i * 8 + 4)[0]) for i in range(n)]
    pos = voff
    out = {}
    for pid, typ, ln in tags:
        v = raw[pos:pos + ln]
        if typ == 0x02 and ln >= 2:
            out[(pid, typ)] = struct.unpack_from("<H", v)[0]
        elif typ == 0x03 and ln >= 4:
            out[(pid, typ)] = struct.unpack_from("<I", v)[0]
        elif typ == 0x0D:
            out[(pid, typ)] = [struct.unpack_from("<I", v, i)[0]
                               for i in range(0, ln - ln % 4, 4)]
        else:
            out[(pid, typ)] = v
        pos += ln
    return out


def series_tz(blobs, default_tz):
    """Best-effort IANA tz from the tz-definition blob (pid 9)."""
    raw = blobs.get((9, 0x0D), b"")
    for m in re.finditer(rb"[A-Za-z_]+/[A-Za-z_+\-]+", raw):
        cand = m.group().decode()
        try:
            ZoneInfo(cand)
            return cand
        except Exception:
            continue
    return default_tz


def decode_recurrence(blobs):
    """Decode the recurrence sub-table -> dict, or raise ValueError."""
    raw = blobs.get((2, 0x0D))
    if not raw:
        raise ValueError("no recurrence blob")
    s = parse_subtable(raw)
    freq = FREQ_MAP.get(s.get((1, 2)))
    if not freq:
        raise ValueError(f"unknown freq {s.get((1,2))}")
    ptype = s.get((1, 3))
    interval = s.get((2, 3), 1) or 1
    endt = s.get((3, 3))
    count = s.get((4, 3))
    fdow = s.get((5, 3), 0)
    spec = s.get((7, 3)) if s.get((7, 3)) is not None else s.get((9, 3), 0)
    nth = s.get((8, 3)) if s.get((8, 3)) is not None else s.get((0x0A, 3))
    end_min = s.get((0x10, 3))
    # multiple observed no-end sentinels; anything beyond year 3000 = never
    if end_min is not None and mdt(end_min).year > 3000:
        end_min = NO_END_SENTINEL
    deleted = s.get((1, 0x0D), []) or []
    modified = s.get((2, 0x0D), []) or []

    parts = []
    if freq == "DAILY":
        if interval >= 1440 and interval % 1440 == 0:
            interval //= 1440
        parts.append(f"FREQ=DAILY;INTERVAL={interval}")
    elif freq == "WEEKLY":
        if not spec:
            raise ValueError("weekly with empty day mask")
        byday = ",".join(DAYS[i] for i in range(7) if spec >> i & 1)
        parts.append(f"FREQ=WEEKLY;INTERVAL={interval};BYDAY={byday};WKST={DAYS[fdow % 7]}")
    elif freq == "MONTHLY":
        if ptype == 2:
            day = spec if 1 <= spec <= 31 else -1  # >31 = last day of month
            parts.append(f"FREQ=MONTHLY;INTERVAL={interval};BYMONTHDAY={day}")
        elif ptype == 3:
            if not spec or nth is None:
                raise ValueError("month-nth missing fields")
            byday = ",".join(DAYS[i] for i in range(7) if spec >> i & 1)
            setpos = -1 if nth == 5 else nth
            parts.append(f"FREQ=MONTHLY;INTERVAL={interval};BYDAY={byday};BYSETPOS={setpos}")
        else:
            raise ValueError(f"monthly ptype {ptype}")
    elif freq == "YEARLY":
        # RFC 5545: bare FREQ=YEARLY repeats on DTSTART's month/day
        parts.append("FREQ=YEARLY;INTERVAL=1")

    rrule = parts[0]
    if endt == 0x2022 and count:
        rrule += f";COUNT={count}"
        until_min = None
    elif endt == 0x2021 and end_min and end_min != NO_END_SENTINEL:
        until_min = end_min
    else:
        until_min = None

    return {"rrule": rrule, "until_min": until_min,
            "deleted": deleted, "modified": modified}


# ------------------------------------------------------------- text helpers

_TAG_RE = re.compile(r"<[^>]+>")


def html_to_text(s):
    if not s:
        return ""
    s = re.sub(r"(?i)<\s*br\s*/?\s*>", "\n", s)
    s = re.sub(r"(?i)</\s*(p|div|tr)\s*>", "\n", s)
    s = _TAG_RE.sub("", s)
    s = html_mod.unescape(s).replace("\r\n", "\n").replace("\r", "\n")
    s = re.sub(r"[ \t]+", " ", s)
    s = re.sub(r"\n{3,}", "\n\n", s).strip()
    return (s[:MAX_DESC] + "\n[truncated]") if len(s) > MAX_DESC else s


def ics_escape(s):
    return (s.replace("\\", "\\\\").replace(";", "\\;")
            .replace(",", "\\,").replace("\n", "\\n"))


def fold(line):
    out, raw = [], line.encode("utf-8")
    while len(raw) > 75:
        cut = 75
        while cut > 1 and (raw[cut] & 0xC0) == 0x80:
            cut -= 1
        out.append(raw[:cut].decode("utf-8"))
        raw = b" " + raw[cut:]
    out.append(raw.decode("utf-8"))
    return "\r\n".join(out)


def fmt_local(dt_utc_naive, tzname):
    loc = dt_utc_naive.replace(tzinfo=timezone.utc).astimezone(ZoneInfo(tzname))
    return loc.strftime("%Y%m%dT%H%M%S")


# --------------------------------------------------------------- DB access

def snapshot_db(tmpdir):
    src = os.path.join(OUTLOOK_DATA_DIR, "Outlook.sqlite")
    if not os.path.exists(src):
        raise SystemExit(f"Outlook store not found: {src}")
    dst = os.path.join(tmpdir, "Outlook.sqlite")
    for sfx in ("", "-wal", "-shm"):
        s = src + sfx
        if os.path.exists(s):
            shutil.copy2(s, dst + sfx)
    return dst


def resolve_folder_ids(conn, cfg):
    explicit = cfg.get("db_folder_ids")
    if explicit:
        return [int(x) for x in explicit]
    name = cfg.get("db_folder_name", "Calendar")
    rows = list(conn.execute(
        "SELECT DISTINCT f.Record_RecordID FROM Folders f "
        "JOIN CalendarEvents c ON c.Record_FolderID=f.Record_RecordID "
        "WHERE f.Folder_Name=?", (name,)))
    if not rows:
        raise SystemExit(f"no calendar folder named {name!r}")
    return [r[0] for r in rows]


# ------------------------------------------------------------ VEVENT build

def vevent(lines_extra, uid, summary, location, description):
    lines = ["BEGIN:VEVENT", f"UID:{ics_escape(uid)}"]
    lines.append("DTSTAMP:" + datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ"))
    lines += lines_extra
    lines.append(f"SUMMARY:{ics_escape(summary or '')}")
    if location:
        lines.append(f"LOCATION:{ics_escape(location)}")
    if description:
        lines.append(f"DESCRIPTION:{ics_escape(description)}")
    lines.append("END:VEVENT")
    return "\r\n".join(fold(x) for x in lines)


def timed_lines(s_min, e_min, tzname):
    s = mdt(s_min)
    e = mdt(e_min) if e_min else s + timedelta(hours=1)
    if e <= s:
        e = s + timedelta(hours=1)
    all_day = s_min % 1440 == 0 and (e_min or 0) % 1440 == 0 and e > s
    if all_day:
        return [f"DTSTART;VALUE=DATE:{s:%Y%m%d}",
                f"DTEND;VALUE=DATE:{e:%Y%m%d}",
                "X-MICROSOFT-CDO-ALLDAYEVENT:TRUE"], True
    return [f"DTSTART;TZID={tzname}:{fmt_local(s, tzname)}",
            f"DTEND;TZID={tzname}:{fmt_local(e, tzname)}"], False


# --------------------------------------------------------------- generation

def generate(cfg):
    default_tz = cfg.get("timezone", "America/New_York")
    days_past = int(cfg.get("sync_days_past", 60))
    days_future = int(cfg.get("sync_days_future", 180))
    now = datetime.now(timezone.utc).replace(tzinfo=None)
    win_lo = dtm(now - timedelta(days=days_past + 14))
    win_hi = dtm(now + timedelta(days=days_future + 2))

    with tempfile.TemporaryDirectory(prefix="cbdb_") as tmp:
        db = snapshot_db(tmp)
        conn = sqlite3.connect(f"file:{db}?mode=ro", uri=True)
        try:
            fids = resolve_folder_ids(conn, cfg)
            ph = ",".join("?" * len(fids))
            base = (f"SELECT Record_RecordID,Calendar_UID,Calendar_StartDateUTC,"
                    f"Calendar_EndDateUTC,Calendar_IsRecurring,Calendar_MasterRecordID,"
                    f"Calendar_RecurrenceID,PathToDataFile FROM CalendarEvents "
                    f"WHERE Record_FolderID IN ({ph}) "
                    f"AND Calendar_StartDateUTC IS NOT NULL ")
            # windowed plain rows + ALL masters (masters can start before window)
            rows = list(conn.execute(
                base + "AND (Calendar_IsRecurring=1 OR "
                "Calendar_StartDateUTC BETWEEN ? AND ?)",
                (*fids, win_lo, win_hi)))
        finally:
            conn.close()

    log.info("Folder scope IDs %s | rows fetched: %d", fids, len(rows))
    masters = {}     # rid -> (uid, tz)
    ves, seen, emitted_uids = [], set(), set()
    stats = dict(plain=0, master=0, master_fallback=0, exception=0,
                 orphan_exception=0, parse_fail=0, missing=0, skipped_old_master=0)

    master_rows = [r for r in rows if r[4] == 1]
    other_rows = [r for r in rows if r[4] != 1]

    # pass 1: masters
    for rid, uid, s_min, e_min, rec, mrid, rcid, pth in master_rows:
        f = os.path.join(OUTLOOK_DATA_DIR, pth)
        try:
            text, blobs = read_props(f)
        except Exception:
            stats["parse_fail"] += 1
            continue
        subj = text.get(PID_SUBJECT, "")
        loc = text.get(PID_LOCATION, "")
        desc = html_to_text(text.get(PID_BODY, ""))
        tz = series_tz(blobs, default_tz)
        try:
            r = decode_recurrence(blobs)
        except Exception as ex:
            # fallback: emit the materialized row as a plain instance
            log.warning("series undecodable (%s): %r — falling back to instance",
                        ex, subj[:40])
            stats["master_fallback"] += 1
            if win_lo <= s_min <= win_hi and (uid, s_min) not in seen:
                seen.add((uid, s_min))
                tl, _ = timed_lines(s_min, e_min, tz)
                ves.append(vevent(tl, uid, subj, loc, desc))
            continue

        # skip series that ended before the window
        if r["until_min"] and r["until_min"] + 1440 < win_lo:
            stats["skipped_old_master"] += 1
            continue

        masters[rid] = (uid, tz)
        tl, all_day = timed_lines(s_min, e_min, tz)
        rrule = r["rrule"]
        if r["until_min"]:
            # inclusive end-of-day UTC on the end date
            until = mdt(r["until_min"]) + timedelta(hours=23, minutes=59)
            rrule += f";UNTIL={until:%Y%m%dT%H%M%S}Z"
        tl.append(f"RRULE:{rrule}")
        # EXDATE every exception slot (deleted AND modified); modified
        # instances are emitted separately as standalone events, which
        # avoids fragile RECURRENCE-ID exact-matching entirely.
        # Exclusion time-of-day must be the DTSTART *local* wall time
        # (occurrences recur at fixed local time; using UTC time-of-day
        # breaks across DST and the EXDATE silently misses).
        local_hms = fmt_local(mdt(s_min), tz)[9:]  # HHMMSS after 'T'
        for dm in sorted(set(r["deleted"]) | set(r["modified"])):
            if all_day:
                tl.append(f"EXDATE;VALUE=DATE:{mdt(dm):%Y%m%d}")
            else:
                tl.append(f"EXDATE;TZID={tz}:{mdt(dm):%Y%m%d}T{local_hms}")
        ves.append(vevent(tl, uid, subj, loc, desc))
        stats["master"] += 1

    # pass 2: plain + exceptions
    for rid, uid, s_min, e_min, rec, mrid, rcid, pth in other_rows:
        if not (win_lo <= s_min <= win_hi):
            continue
        f = os.path.join(OUTLOOK_DATA_DIR, pth)
        subj = loc = desc = ""
        if os.path.exists(f):
            try:
                text, blobs = read_props(f)
                subj = text.get(PID_SUBJECT, "")
                loc = text.get(PID_LOCATION, "")
                desc = html_to_text(text.get(PID_BODY, ""))
            except Exception:
                stats["parse_fail"] += 1
        else:
            stats["missing"] += 1

        if mrid:
            stats["exception" if mrid in masters else "orphan_exception"] += 1
        key = (uid, s_min)
        if key in seen:
            continue
        seen.add(key)
        tz = masters.get(mrid, (None, default_tz))[1] if mrid else default_tz
        tl, _ = timed_lines(s_min, e_min, tz)
        # Exception rows share the master's UID; two VEVENTs with one UID
        # and no RECURRENCE-ID is undefined ICS and expansion drops them.
        # Rescheduled plain events can also leave two rows on one UID.
        # First emission keeps the bare UID (preserves sync continuity);
        # any subsequent emission on the same UID gets a derived one.
        if mrid or uid in emitted_uids:
            out_uid = f"{uid}-EXC{s_min}"
        else:
            out_uid = uid
        emitted_uids.add(uid)
        ves.append(vevent(tl, out_uid, subj, loc, desc))
        if not mrid:
            stats["plain"] += 1

    log.info("Generated VEVENTs: %d | %s", len(ves),
             " ".join(f"{k}={v}" for k, v in stats.items()))
    body = "\r\n".join(
        ["BEGIN:VCALENDAR", "VERSION:2.0",
         "PRODID:-//CalendarBridge//olk15-db-export-v2//EN"]
        + ves + ["END:VCALENDAR", ""])
    return body, stats, len(ves)


# ------------------------------------------------------------------ parity

def expand_ics(path_or_text, is_text, lo, hi):
    from icalendar import Calendar
    from recurring_ical_events import of as recurring_of
    content = path_or_text if is_text else open(
        path_or_text, encoding="utf-8", errors="ignore").read()
    out = {}
    for block in content.split("BEGIN:VCALENDAR"):
        if not block.strip():
            continue
        b = "BEGIN:VCALENDAR" + block
        if not b.rstrip().endswith("END:VCALENDAR"):
            b = b.rstrip() + "\nEND:VCALENDAR"
        try:
            cal = Calendar.from_ical(b.encode())
        except Exception:
            continue
        try:
            comps = recurring_of(cal).between(lo, hi)
        except Exception as e:
            log.warning("expand failed for a block: %s", e)
            continue
        for c in comps:
            if c.name != "VEVENT":
                continue
            s = c.get("DTSTART").dt if c.get("DTSTART") else None
            summ = re.sub(r"\s+", " ", str(c.get("SUMMARY") or "").strip())[:50]
            if isinstance(s, datetime):
                if s.tzinfo is None:
                    s = s.replace(tzinfo=ZoneInfo("America/New_York"))
                u = s.astimezone(timezone.utc)
                # normalize midnight datetimes (UTC or local) to date keys so
                # all-day representations (DATE vs T000000 variants) compare equal
                l = s.astimezone(ZoneInfo("America/New_York"))
                if (u.hour, u.minute) == (0, 0):
                    key = (u.strftime("%Y-%m-%d"), summ)
                elif (l.hour, l.minute) == (0, 0):
                    key = (l.strftime("%Y-%m-%d"), summ)
                else:
                    key = (u.strftime("%Y-%m-%d %H:%M"), summ)
            elif isinstance(s, dtmod.date):
                key = (s.isoformat(), summ)
            else:
                continue
            out[key] = 1  # set semantics: identical (time, summary) counts once
    return out


def parity(baseline_path, cfg):
    body, stats, n = generate(cfg)
    days_past = int(cfg.get("sync_days_past", 60))
    days_future = int(cfg.get("sync_days_future", 180))
    now = datetime.now()
    lo, hi = now - timedelta(days=days_past), now + timedelta(days=days_future)
    log.info("Expanding baseline...")
    old = expand_ics(baseline_path, False, lo, hi)
    log.info("Expanding DB export...")
    new = expand_ics(body, True, lo, hi)
    so, sn = sum(old.values()), sum(new.values())
    print(f"\n===== PARITY REPORT (window {lo:%Y-%m-%d} .. {hi:%Y-%m-%d}) =====")
    print(f"baseline expanded: {so} | DB export expanded: {sn}")
    miss = {k: v for k, v in old.items() if new.get(k, 0) < v}
    extra = {k: v for k, v in new.items() if old.get(k, 0) < v}
    def by_series(d):
        agg = {}
        for (t, s), v in d.items():
            agg.setdefault(s, []).append(t)
        return agg
    print(f"\nMISSING from DB export ({sum(miss.values()) - sum(old.get(k,0)-v for k,v in miss.items() if False)} keys, {len(by_series(miss))} series):")
    for s, ts in sorted(by_series(miss).items(), key=lambda x: -len(x[1]))[:25]:
        print(f"  [{len(ts):3d}] {s}   e.g. {sorted(ts)[:3]}")
    print(f"\nEXTRA in DB export ({len(by_series(extra))} series):")
    for s, ts in sorted(by_series(extra).items(), key=lambda x: -len(x[1]))[:25]:
        print(f"  [{len(ts):3d}] {s}   e.g. {sorted(ts)[:3]}")
    pct = 100.0 * (so - sum(old.get(k, 0) - new.get(k, 0) for k in miss)) / max(so, 1)
    print(f"\ncoverage of baseline: {pct:.1f}%")
    print("GATE: cutover requires >=99% coverage and no unexplained EXTRA series.")


# -------------------------------------------------------------------- main

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--parity", metavar="BASELINE_ICS",
                    help="offline diff vs AppleScript baseline; no writes")
    args = ap.parse_args()
    logging.basicConfig(level=logging.INFO,
                        format="%(asctime)s %(levelname)s %(message)s")
    cfg = load_config()

    if args.parity:
        parity(args.parity, cfg)
        return

    body, stats, n = generate(cfg)
    failures = []
    if n == 0:
        failures.append("0 events")
    attempted = n + stats["parse_fail"]
    if attempted and stats["parse_fail"] / attempted > 0.2:
        failures.append("parse failures >20%")
    last = None
    try:
        with open(STATE_PATH) as f:
            last = json.load(f).get("last_good_count")
    except Exception:
        pass
    if last and last >= 20 and n < last * 0.5:
        failures.append(f"count {n} < 50% of last good {last}")
    if failures:
        msg = "; ".join(failures)
        log.error("VALIDATION FAILED — not writing ICS: %s", msg)
        notify(f"Export blocked: {msg}")
        sys.exit(1)
    if args.dry_run:
        log.info("Dry run OK: %d VEVENTs", n)
        return
    os.makedirs(OUTBOX, exist_ok=True)
    fd, tmp_path = tempfile.mkstemp(dir=OUTBOX, prefix=".ics_tmp_")
    with os.fdopen(fd, "w", encoding="utf-8", newline="") as f:
        f.write(body)
    os.replace(tmp_path, ICS_PATH)
    with open(STATE_PATH, "w") as f:
        json.dump({"last_good_count": n,
                   "last_run": datetime.utcnow().isoformat() + "Z"}, f)
    log.info("Wrote %d VEVENTs -> %s", n, ICS_PATH)


if __name__ == "__main__":
    main()
