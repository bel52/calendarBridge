-- exportEvents.scpt – Outlook → ~/calendarBridge/outbox/ (.ics)
-- Exports all recurring series (regardless of start date),
-- and non-recurring events only if their start is within the window.

on run
    -- ===== CONFIG =====
    set targetCalIndex to 2 -- adjust if needed
    set exportDaysBack to 60
    set exportDaysAhead to 120
    -- ===================

    -- Paths
    set outboxFolder to (path to home folder as text) & "calendarBridge:outbox:"
    set posixOutbox to POSIX path of outboxFolder
    set quarantineFile to POSIX path of ((path to home folder as text) & "calendarBridge:quarantine.txt")

    do shell script "mkdir -p " & quoted form of posixOutbox
    do shell script "rm -f " & quoted form of posixOutbox & "*.ics" -- wildcard outside quotes
    do shell script "touch " & quoted form of quarantineFile

    -- Date window (outside tell; reference with 'my')
    set startDate to (current date) - (exportDaysBack * days)
    set endDate to (current date) + (exportDaysAhead * days)

    tell application "Microsoft Outlook"
        activate
        set targetCal to calendar targetCalIndex
        set allEvents to calendar events of targetCal

        set exportedCount to 0
        set skippedCount to 0

        repeat with ev in allEvents
            try
                -- detect recurrence
                set isRecurring to false
                try
                    set r to recurrence of ev
                    if r is not missing value then set isRecurring to true
                end try

                set shouldExport to false
                if isRecurring then
                    set shouldExport to true
                else
                    set st to start time of ev
                    if (st is greater than or equal to my startDate) and (st is less than or equal to my endDate) then
                        set shouldExport to true
                    end if
                end if

                if shouldExport then
                    set uidStr to (id of ev) as string
                    set eventDetails to icalendar data of ev

                    set savePath to (my posixOutbox) & uidStr & ".ics"
                    set fileRef to open for access (POSIX file savePath) with write permission
                    set eof of fileRef to 0
                    write eventDetails to fileRef
                    close access fileRef

                    set exportedCount to exportedCount + 1
                    delay 0.05
                end if
            on error errm
                try
                    set uidStr to (id of ev) as string
                    do shell script "echo " & quoted form of uidStr & " >> " & quoted form of (my quarantineFile)
                    set skippedCount to skippedCount + 1
                end try
            end try
        end repeat

        log "Exported " & exportedCount & " events, skipped " & skippedCount
    end tell
end run
