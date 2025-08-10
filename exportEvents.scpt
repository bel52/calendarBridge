-- exportEvents.scpt – Outlook → ~/calendarBridge/outbox/ (.ics)
-- Minimal & robust: iterate events and filter by start time inline.
-- No Unicode operators, no 'whose' filters, no tmp vars like 'st'.

on run
    -- ===== CONFIG =====
    set targetCalIndex to 2 -- change if needed (your list shows: 1|Calendar, 2|Calendar, 3|Birthdays)
    set exportDaysBack to 60
    set exportDaysAhead to 120
    -- ===================

    -- Paths
    set outboxFolder to (path to home folder as text) & "calendarBridge:outbox:"
    set posixOutbox to POSIX path of outboxFolder
    set quarantineFile to POSIX path of ((path to home folder as text) & "calendarBridge:quarantine.txt")

    do shell script "mkdir -p " & quoted form of posixOutbox
    -- wildcard must live outside quotes:
    do shell script "rm -f " & quoted form of posixOutbox & "*.ics"
    do shell script "touch " & quoted form of quarantineFile

    -- Date window (define OUTSIDE tell; reference with 'my' inside)
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
                if ((start time of ev) is greater than or equal to my startDate) and ((start time of ev) is less than or equal to my endDate) then
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
