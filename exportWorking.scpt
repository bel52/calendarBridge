tell application "Microsoft Outlook"
    set cal_list to every calendar whose name is "Calendar"
    set my_calendar to item 2 of cal_list
    set output_file to ((path to home folder) as text) & "calendarBridge:outbox:calendar.ics"
    
    do shell script "mkdir -p " & quoted form of (POSIX path of (path to home folder) & "calendarBridge/outbox/")
    
    export my_calendar in file output_file as ics
end tell
