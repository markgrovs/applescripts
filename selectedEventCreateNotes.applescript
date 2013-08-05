get_selected_event()

on get_selected_event()
	tell application "Calendar" to activate
	delay 2
	tell application "System Events"
		tell process "Calendar"
			keystroke return
			keystroke "c" using {command down}
			keystroke return
		end tell
	end tell
	
	tell application "Calendar"
		activate
		
		try
			set myClipboard to the clipboard
			display dialog "Event's UD is: " & first paragraph of myClipboard
			set eventName to first paragraph of myClipboard
			--set defcalUID to first paragraph of (do shell script "defaults read com.apple.iCal CalDefaultCalendar")
			set defCal to calendar "Calendar"
			
			set myEvents to events of defCal whose summary is eventName
			get count of  myEvents
			repeat with myEvent in myEvents
				display dialog "ReCRR: " & recurrence of myEvent
				get start date of myEvent
				repeat with att in attendees of myEvent
					get display name of att
				end repeat
			end repeat
			(*
			if (recurance of myEvent not empty ) then
				set recEvent to recurance of myEvent
				display dialog "Rec: " & start date of recEvent
			end if
			*)
			
			--display dialog "Event: " & attendees of myEvent
		end try
	end tell
end get_selected_event











