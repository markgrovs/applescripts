on run {}
	tell application "Microsoft Outlook"
		set selectedMessages to current messages
		
		if selectedMessages is {} then
			display dialog "Please select a message first" with icon 1
			return
		end if
		
		set archiveFolder to mail folders whose name is "Archive"
		if ((count of archiveFolder) > 1) then
			set archiveFolder to (item 1 of archiveFolder)
		end if
		
		repeat with theMessage in selectedMessages
			if archiveFolder is not missing value then
				move theMessage to archiveFolder
			end if
		end repeat
	end tell
end run
