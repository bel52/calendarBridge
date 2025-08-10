tell application "Microsoft Outlook"
    -- Get the second Calendar
    set calList to calendars whose name is "Calendar"
    set targetCal to item 2 of calList
    
    -- Set export path
    set exportPath to (path to home folder as text) & "calendarBridge:full_calendar.ics"
    
    -- Export the entire calendar
    export targetCal in file exportPath as ics
    
    return "Calendar exported"
end tell
