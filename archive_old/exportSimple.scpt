tell application "Microsoft Outlook"
    -- Get calendars named "Calendar"
    set calList to calendars whose name is "Calendar"
    set targetCal to item 2 of calList
    
    -- Export path
    set desktopPath to (path to desktop as text) & "outlook_export.ics"
    
    -- Export the calendar
    export targetCal to file desktopPath as "ics"
    
    return "Exported calendar to Desktop"
end tell
