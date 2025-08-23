-- ============================================================================
-- AppleScript to Trigger a Bulk Export of a Named Outlook Calendar by Index
-- Version: 9.0 (Handles duplicate names)
-- ============================================================================

on run argv
    if (count of argv) < 2 then
        error "Expected 2 arguments: calendar name and index."
    end if
    set calendarNameToFind to item 1 of argv
    set calendarIndexToUse to (item 2 of argv) as integer

    set exportFolder to (system attribute "HOME") & "/calendarBridge/outbox/"
    set exportFile to "outlook_full_export.ics"
    set exportPath to exportFolder & exportFile

    do shell script "mkdir -p " & quoted form of exportFolder
    do shell script "rm -f " & quoted form of exportPath

    tell application "Microsoft Outlook"
        set foundCalendar to missing value
        try
            -- Find all calendars matching the name
            set matchingCalendars to (every calendar whose name is calendarNameToFind)
            
            if (count of matchingCalendars) < calendarIndexToUse then
                error "Found " & (count of matchingCalendars) & " calendar(s) named '" & calendarNameToFind & "', but you requested index " & calendarIndexToUse & "."
            end if
            
            -- Select the calendar at the specified index
            set foundCalendar to item calendarIndexToUse of matchingCalendars
            
        on error errNum
            error "Could not find calendar '" & calendarNameToFind & "' at index " & calendarIndexToUse & ". Error: " & errNum
        end try

        if foundCalendar is not missing value then
            try
                save foundCalendar in (POSIX file exportPath)
            on error errMsg number errNum
                error "Outlook failed to save the calendar. Error (" & errNum & "): " & errMsg
            end try
        else
            error "Failed to get a reference to the calendar."
        end if
    end tell
    log "Successfully exported calendar '" & calendarNameToFind & "' (index " & calendarIndexToUse & ") to: " & exportPath
end run
