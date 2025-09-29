// outlook_write_events.js
// Usage:
//   osascript -l JavaScript outlook_write_events.js /tmp/public_events.json "Calendar 2" "Imported: Detroit Lions"
//
// Requirements:
// - Use the **legacy/classic Outlook** (AppleScript/JXA-enabled). The "New Outlook" app
//   largely removed scripting support. If you're on New Outlook, toggle it off (Help -> Revert to Legacy Outlook).
//
// Behavior:
// - Ensures a calendar named targetCalendarName exists (creates if needed).
// - Ensures categoryName exists (creates if needed).
// - For each event from JSON, searches for an existing event in that calendar by a hidden
//   "[ICSUID: ...]" tag embedded in the notes. If found, updates it; else creates it.
// - Times: accepts ISO 8601; if string ends with 'Z' it's UTC, else treated as local.
//
// IMPORTANT: Outlook JXA object model can vary by version. This script targets Office 16+ (legacy).
//
function run(argv) {
  if (argv.length < 3) {
    console.log("Usage: osascript -l JavaScript outlook_write_events.js <jsonPath> <calendarName> <categoryName>");
    return;
  }
  const jsonPath = argv[0];
  const calendarName = argv[1];
  const categoryName = argv[2];

  const app = Application('Microsoft Outlook');
  app.includeStandardAdditions = true;

  // Read JSON
  const jsonText = app.doShellScript(`/bin/cat ${escapePath(jsonPath)}`);
  const data = JSON.parse(jsonText);
  const events = (data && data.events) || [];

  // Get default account & root
  const accounts = app.accounts();
  if (accounts.length === 0) {
    throw new Error("No Outlook accounts found.");
  }

  // Find (or create) target calendar under the *On My Computer* or first account calendar set.
  // We search all calendars and pick the one matching name exactly.
  const allCalendars = [];
  accounts.forEach(acc => {
    try {
      // top-level calendar folders for this account
      acc.calendars().forEach(c => allCalendars.push(c));
    } catch (e) {}
    try {
      // "On My Computer" calendars might be under application.calendars()
      // (depends on Outlook version). We'll gather those too.
      app.calendars().forEach(c => allCalendars.push(c));
    } catch (e) {}
  });
  // Deduplicate by id
  const seen = {};
  const calendars = allCalendars.filter(c => {
    const id = (c.id && String(c.id())) || String(c);
    if (seen[id]) return false;
    seen[id] = true;
    return true;
  });

  let targetCal = calendars.find(c => (c.name && c.name()) === calendarName);
  if (!targetCal) {
    // Try to create under the first account
    const parent = accounts[0];
    targetCal = app.Calendar({name: calendarName});
    try {
      parent.make({new: targetCal, at: parent});
    } catch (e) {
      // Fallback: create at application level
      app.make({new: targetCal, at: app});
    }
    console.log(`Created calendar "${calendarName}"`);
  }

  // Category ensure
  let targetCategory = ensureCategory(app, categoryName);

  // Build a map of existing events that have our ICSUID marker (quick lookup)
  // We’ll scan a broad range (±365 days from today) to catch seasonal schedules.
  const now = new Date();
  const startScan = new Date(now.getTime() - 365*24*3600*1000);
  const endScan   = new Date(now.getTime() + 365*24*3600*1000);

  // Fetch events in range from targetCal (Outlook JXA model: calendarEvents whose start time within range)
  const existing = eventsInRange(app, targetCal, startScan, endScan);
  // map by our embedded UID
  const byUid = {};
  existing.forEach(ev => {
    try {
      const body = (ev.content && ev.content()) || "";
      const m = body.match(/\[ICSUID:\s*([^\]]+)\]/);
      if (m) {
        byUid[m[1]] = ev;
      }
    } catch (e) {}
  });

  let created = 0, updated = 0, skipped = 0;

  events.forEach(item => {
    const uid = String(item.uid || '').trim();
    if (!uid) { skipped++; return; }

    const start = parseIso(item.start);
    const end = item.end ? parseIso(item.end) : new Date(start.getTime() + 60*60*1000);
    const allDay = !!item.all_day;

    // Compose body with hidden UID tag for future sync
    const details = (item.description || '').trim();
    const body = details
      ? details + "\n\n[ICSUID: " + uid + "]"
      : "[ICSUID: " + uid + "]";

    let ev = byUid[uid];
    if (ev) {
      // Update existing
      try {
        if (item.summary) ev.subject = item.summary;
        if (item.location) ev.location = item.location;

        if (allDay) {
          ev.allDayEvent = true;
          // Outlook typically uses start at midnight, end next midnight for all-day
          const s = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 0,0,0);
          const e = new Date(s.getTime() + 24*3600*1000);
          ev.startTime = s;
          ev.endTime = e;
        } else {
          ev.allDayEvent = false;
          ev.startTime = start;
          ev.endTime = end;
        }
        ev.content = body;

        // Category
        try { ev.category = targetCategory; } catch (e) {}

        updated++;
      } catch (e) {
        console.log("Update failed for UID " + uid + ": " + e);
        skipped++;
      }
    } else {
      // Create new
      try {
        const newEv = app.CalendarEvent({
          subject: item.summary || "(No title)",
          location: item.location || "",
          content: body,
          calendar: targetCal
        });

        app.make({ new: newEv, at: targetCal });

        if (allDay) {
          const s = new Date(start.getFullYear(), start.getMonth(), start.getDate(), 0,0,0);
          const e = new Date(s.getTime() + 24*3600*1000);
          newEv.allDayEvent = true;
          newEv.startTime = s;
          newEv.endTime = e;
        } else {
          newEv.allDayEvent = false;
          newEv.startTime = start;
          newEv.endTime = end;
        }

        try { newEv.category = targetCategory; } catch (e) {}
        created++;
      } catch (e) {
        console.log("Create failed for UID " + uid + ": " + e);
        skipped++;
      }
    }
  });

  console.log(`Done. Created: ${created}, Updated: ${updated}, Skipped: ${skipped}`);
}

function escapePath(p) {
  // crude shell escaping for paths
  return "'" + String(p).replace(/'/g, "'\\''") + "'";
}

function ensureCategory(app, name) {
  try {
    const cats = app.categories();
    const hit = cats.find(c => (c.name && c.name()) === name);
    if (hit) return hit;
    const cat = app.Category({name});
    app.make({new: cat, at: app});
    return cat;
  } catch (e) {
    // Older Outlooks: category may not be settable; return null
    return null;
  }
}

function eventsInRange(app, calendar, startDate, endDate) {
  // Some Outlook builds support calendar.calendarEvents(); others need application-level filtering.
  // We'll try both; fall back to scanning all and filtering in JS if needed.
  let list = [];
  try {
    list = calendar.calendarEvents().filter(ev => {
      const s = safeDate(ev.startTime);
      return s && s >= startDate && s <= endDate;
    });
  } catch (e) {
    try {
      list = app.calendarEvents().filter(ev => {
        const cal = ev.calendar && ev.calendar();
        const s = safeDate(ev.startTime);
        return cal && (cal.id && cal.id() === calendar.id()) && s && s >= startDate && s <= endDate;
      });
    } catch (e2) {
      list = []; // best effort
    }
  }
  return list;
}

function safeDate(prop) {
  try { return prop(); } catch (e) { return null; }
}

function parseIso(str) {
  // Accepts 'YYYY-MM-DDTHH:MM:SS' (local) or '...Z' (UTC)
  if (str.endsWith('Z')) {
    return new Date(str); // JS Date parses Z as UTC
  }
  // treat as local
  return new Date(str);
}
