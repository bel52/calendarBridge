-- AppleScript: diagnose_calendars.scpt
-- Lists all calendars accessible via Outlook AppleScript and counts exportable events

on run
    set exportDaysBack to 7
    set exportDaysAhead to 120

    set startDate to (current date) - (exportDaysBack * days)
    set endDate to (current date) + (exportDaysAhead * days)

    set summaryText to "🧪 Calendar Diagnostic Summary:\n\n"

    tell application "Microsoft Outlook"
        try
            set calList to calendars
            set calendarCount to count of calList
            if calendarCount = 0 then
                display dialog "❌ No calendars found in Outlook." buttons {"OK"} default button "OK"
                return
            end if

            repeat with calIndex from 1 to calendarCount
                set thisCal to item calIndex of calList
                set calName to name of thisCal
                try
                    set eventsInRange to calendar events of thisCal whose start time ≥ startDate and start time ≤ endDate
                    set eventCount to count of eventsInRange
                on error errCal
                    set eventCount to "ERROR (" & errCal & ")"
                end try
                set summaryText to summaryText & calName & ": " & eventCount & " events\n"
            end repeat

            display dialog summaryText buttons {"OK"} default button "OK"
            return summaryText

        on error errAll
            display dialog "❗️Fatal calendar scan error: " & errAll buttons {"OK"} default button "OK"
            return "Error: " & errAll
        end try
    end tell
end run
