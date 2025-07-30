-- AppleScript: exportEvents.scpt  (2025-07-29 window-aware)
-- Exports every Outlook calendar the Mac client can access,
-- but only events that INTERSECT a 127-day sliding window
--   ( 7 days back  ⟶  120 days ahead ).

on run
    -- ─── Paths ────────────────────────────────────────────────
    set homePath to path to home folder as text
    set outboxPath to homePath & "calendarBridge:outbox:"
    set posixOutbox to POSIX path of (path to home folder) & "calendarBridge/outbox/"

    -- ─── Window definition ────────────────────────────────────
    set exportDaysBack to 7
    set exportDaysAhead to 120
    set startDate to (current date) - (exportDaysBack * days)
    set endDate   to (current date) + (exportDaysAhead * days)

    -- ─── Misc config ──────────────────────────────────────────
    set skipUIDs to {"5471", "5472"}

    do shell script "mkdir -p " & quoted form of posixOutbox

    tell application "Microsoft Outlook"
        try
            set calList to calendars
            if (count of calList) = 0 then error "❌ No calendars found in Outlook."

            set totalExported to 0
            set totalSkipped  to 0
            set totalErrors   to 0
            set globalEventIndex to 0

            repeat with calIndex from 1 to count of calList
                set thisCal to item calIndex of calList
                set calName to name of thisCal

                -- Include ONLY events that overlap the window
                set eventsInRange to calendar events of thisCal ¬
                    whose start time ≤ endDate and end time ≥ startDate

                repeat with evt in eventsInRange
                    set globalEventIndex to globalEventIndex + 1
                    try
                        set evtID to id of evt as string
                        if skipUIDs contains evtID then
                            set totalSkipped to totalSkipped + 1
                        else
                            set safeUID to do shell script "echo " & quoted form of evtID & " | tr -cd '[:alnum:]_-.@'"
                            if safeUID is "" then set safeUID to "cal" & calIndex & "_event_" & globalEventIndex
                            set fileName to calName & "_" & safeUID & ".ics"
                            set filePath to outboxPath & fileName

                            try
                                save evt in file filePath as "ics"
                                set totalExported to totalExported + 1
                            on error
                                try
                                    save evt in file (outboxPath & "event_" & globalEventIndex & ".ics") as "ics"
                                    set totalExported to totalExported + 1
                                on error
                                    set totalErrors to totalErrors + 1
                                end try
                            end try
                        end if
                    on error
                        set totalErrors to totalErrors + 1
                    end try
                end repeat
            end repeat

            set msg to "✅ Export complete. Exported " & totalExported & ¬
                       ", skipped " & totalSkipped & ", errors: " & totalErrors
            display notification msg with title "Calendar Bridge"
            return msg

        on error errAll
            display dialog "❗️Fatal export error: " & errAll buttons {"OK"} default button "OK"
            return "Error: " & errAll
        end try
    end tell
end run
