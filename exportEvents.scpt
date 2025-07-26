-- exportEvents.scpt – Outlook → ~/calendarBridge/outbox/ (.ics)
-- • Deletes stale *.ics correctly          (rm line fixed)
-- • Tiny 0.1-s delay per save (prevents -2700)
-- • Logs bad UIDs to quarantine.txt

on run
    -- CONFIG --------------------------------------------------------------
    set targetCalName  to "Calendar"
    set targetCalIndex to 2
    set exportDaysBack to 7
    set exportDaysAhead to 120
    -----------------------------------------------------------------------

    -- Paths ---------------------------------------------------------------
    set outboxFolder   to (path to home folder as text) & "calendarBridge:outbox:"
    set posixOutbox    to POSIX path of outboxFolder
    set quarantineFile to POSIX path of ((path to home folder as text) & "calendarBridge:quarantine.txt")

    do shell script "mkdir -p " & quoted form of posixOutbox
    -- ✨  wild-card OUTSIDE the quotes
    do shell script "rm -f " & quoted form of posixOutbox & "*.ics"
    do shell script "touch " & quoted form of quarantineFile

    -- Date window ---------------------------------------------------------
    set startDate to (current date) - (exportDaysBack * days)
    set endDate   to (current date) + (exportDaysAhead * days)

    tell application "Microsoft Outlook"
        activate
        set cals to every calendar whose name is targetCalName
        if (count of cals) < targetCalIndex then error "Calendar index not found"
        set targetCal to item targetCalIndex of cals

        set evtsInRange to (calendar events of targetCal ¬
            whose start time ≥ startDate and start time ≤ endDate)

        set exportedCount to 0
        set skippedCount  to 0

        repeat with ev in evtsInRange
            set uidStr to id of ev as string
            try
                set savePath to posixOutbox & uidStr & ".ics"
                save ev in (POSIX file savePath) as "ics"
                set exportedCount to exportedCount + 1
                delay 0.1
            on error errm
                log "❌ Skipped UID " & uidStr & " : " & errm
                do shell script "echo " & quoted form of uidStr & " >> " & quoted form of quarantineFile
                set skippedCount to skippedCount + 1
            end try
        end repeat
    end tell

    log "✓ Exported " & exportedCount & " events, skipped " & skippedCount
end run
