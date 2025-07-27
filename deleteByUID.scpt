-- ⚠️ Deletes events by UID. Use carefully.
set uidsToDelete to {"5471", "5472"}
set deletedCount to 0
set deletedUIDs to {}

tell application "Microsoft Outlook"
	set allCalendars to calendars
	repeat with cal in allCalendars
		repeat with targetUID in uidsToDelete
			try
				set matches to every calendar event of cal whose id as string is targetUID
				repeat with ev in matches
					set deletedCount to deletedCount + 1
					set end of deletedUIDs to targetUID
					delete ev
				end repeat
			end try
		end repeat
	end repeat
end tell

-- Show results
if deletedCount > 0 then
	display dialog "✅ Deleted " & deletedCount & " events: " & deletedUIDs as string buttons {"OK"} default button "OK"
else
	display dialog "⚠️ No matching events found to delete." buttons {"OK"} default button "OK"
end if
