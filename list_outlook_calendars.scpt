-- Prints index and name of each Outlook calendar
on run
    set outLines to {}
    tell application "Microsoft Outlook"
        activate
        set calList to every calendar
        set n to count of calList
        repeat with i from 1 to n
            set calRef to item i of calList
            set calName to (name of calRef) as string
            set end of outLines to (i as string) & " | " & calName
        end repeat
    end tell
    set AppleScript's text item delimiters to linefeed
    return outLines as text
end run
