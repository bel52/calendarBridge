#!/bin/bash
echo "🧪 Testing Calendar Bridge Setup"
echo "================================"

cd ~/calendarBridge

# Test 1: Check Outlook calendars
echo -e "\n1️⃣ Checking Outlook calendars:"
osascript -e 'tell application "Microsoft Outlook" to get name of every calendar'

# Test 2: Count events in main calendar
echo -e "\n2️⃣ Counting total events in first Calendar:"
osascript -e 'tell application "Microsoft Outlook"
    set cal to first calendar whose name is "Calendar"
    return "Total events: " & (count of calendar events of cal)
end tell'

# Test 3: Count future events
echo -e "\n3️⃣ Counting future events:"
osascript -e 'tell application "Microsoft Outlook"
    set cal to first calendar whose name is "Calendar"
    set futureEvents to calendar events of cal whose start time > (current date)
    return "Future events: " & (count of futureEvents)
end tell'

# Test 4: Show next 5 events
echo -e "\n4️⃣ Next 5 events:"
osascript -e 'tell application "Microsoft Outlook"
    set cal to first calendar whose name is "Calendar"
    set futureEvents to calendar events of cal whose start time > (current date)
    set output to ""
    repeat with i from 1 to 5
        if i ≤ (count of futureEvents) then
            set evt to item i of futureEvents
            set output to output & (subject of evt) & " - " & (start time of evt as string) & "\n"
        end if
    end repeat
    return output
end tell'

# Test 5: Try export
echo -e "\n5️⃣ Testing export:"
osascript exportEvents.scpt

# Test 6: Check exported files
echo -e "\n6️⃣ Exported files:"
ls -la outbox/*.ics 2>/dev/null | wc -l
echo "files found in outbox/"

echo -e "\n✅ Test complete!"
