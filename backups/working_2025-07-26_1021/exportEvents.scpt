-- exportEvents.scpt – Outlook → ~/calendarBridge/outbox/ (.ics)

on run
    -- CONFIG
    set targetCalName to "Calendar"
    set targetCalIndex to 2
    set exportDaysBack to 7
    set exportDaysAhead to 120
    set skipUIDs to {"5471", "5472"}

    -- Prepare outbox
    set outboxFolder to (path to home folder as text) & "calendarBridge:outbox:"
    set posixOutbox to POSIX path of outboxFolder
    do shell script "mkdir -p " & quoted form of posixOutbox
    do shell script "rm -f " & quoted form of (posixOutbox & "*.ics")

    -- Date window
    set startDate to (current date) - (exportDaysBack * days)
    set endDate to (current date) + (exportDaysAhead * days)

    tell application "Microsoft Outlook"
        activate

        -- Find calendar
        set cals to every calendar whose name is targetCalName
        if (count of cals) < targetCalIndex then
            log "🚫 Need at least " & targetCalIndex & " calendars named 'Calendar'; found " & (count of cals)
            error number -128
        end if
        set targetCal to item targetCalIndex of cals

        -- Export events
        set exportedCount to 0
        set skippedCount to 0
        repeat with ev in (calendar events of targetCal)
            set evStart to start time of ev
            if evStart is greater than or equal to startDate and evStart is less than or equal to endDate then
                set uidStr to id of ev as string
                if skipUIDs contains uidStr then
                    set skippedCount to skippedCount + 1
                else
                    try
                        set savePath to posixOutbox & uidStr & ".ics"
                        save ev in (POSIX file savePath) as "ics"
                        set exportedCount to exportedCount + 1
                    on error errm
                        log "❌ Skipped UID " & uidStr & " : " & errm
                        set skippedCount to skippedCount + 1
                    end try
                end if
            end if
        end repeat
    end tell

    log "✓ Exported " & exportedCount & " events, skipped " & skippedCount
    return 0
end run
