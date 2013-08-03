tell application "Microsoft Outlook"
	
	set today to (current date) - (24 * 60 * 60)
	set time of today to 0
	log (today)
	
	
	set tomorrow to (today) + (24 * 60 * 60)
	log (tomorrow)
	(*
	set allMeetings to calendar events
	
	repeat with aMeeting in allMeetings
		if is occurrence of aMeeting is true then
			log "All Meeting Title: " & (get subject of aMeeting)
			set rec to recurrence of aMeeting
			log "Date: " & (get start date of rec)
		end if
		
	end repeat
	*)
	set todaysMeetings to calendar events where start time is greater than today and start time is less than tomorrow
	if ((count of todaysMeetings) < 1) then return
	
	repeat with meeting in todaysMeetings
		
		set meetingSubject to get subject of meeting
		set meetingStartTime to get start time of meeting
		set people to get attendees of meeting
		
		set attendeeHeader to "attendees: ["
		if ((count of people) > 0) then
			set lastEmail to email address of the last item of people
		end if
		
		repeat with per in people
			set ea to get email address of per
			set n to get name of ea
			if (ea is not lastEmail) then
				set attendeeHeader to attendeeHeader & n & ", "
			else
				set attendeeHeader to attendeeHeader & n
			end if
		end repeat
		set attendeeHeader to attendeeHeader & "]"
		
		set fileName to my set_file_name(meetingStartTime, meetingSubject, ".md")
		set fileName to my clean_filename(fileName)
		
		set lowerSubject to (my change_case(meetingSubject, 0))
		set lowerSubject to my replace_chars(lowerSubject, " ", "-")
		set filePath to (path to desktop as text) & fileName
		set fileContent to ("---" & return) & ("meeting: " & meetingSubject & return) & ("date: " & my format_date(meetingStartTime, "<<Month 2>>-<<Day 2>>-<<Year>> <<Time>>") & return) & (attendeeHeader & return) & ("---" & return)
		my write_to_file(fileContent, filePath, false)
		
	end repeat
	
end tell

on create_yaml_header(meetingSubject, meetingDate, attendees, location)
	set yaml_delim to "---" & return
	set yaml_header to yaml_delim
	set yaml_header to yaml_header & "meeting: " & meetingSubject & return
	set yaml_header to yaml_header & "date: " & meetingDate & return
	set yaml_header to yaml_header & "attendees: " & attendees & return
	set yaml_header to yaml_header & yaml_delim & return
	return yaml_header
end create_yaml_header

on set_file_name(fdate, subject, extension)
	set fileDate to format_date(fdate, "<<Month 2>>-<<Day 2>>-<<Year>>-")
	return (fileDate & change_case(replace_chars(subject, " ", "-"), 0) & extension)
end set_file_name

on write_to_file(this_data, target_file, append_data)
	try
		set the target_file to the target_file as string
		log (target_file)
		set the open_target_file to open for access file target_file with write permission
		if append_data is false then set eof of the open_target_file to 0
		
		write this_data to the open_target_file starting at eof
		
		close access the open_target_file
		
		return true
	on error errormsg
		log (errormsg)
		try
			close access file target_file
		end try
		return false
	end try
end write_to_file

on change_case(this_text, this_case)
	if this_case is 0 then
		set the comparison_string to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
		set the source_string to "abcdefghijklmnopqrstuvwxyz"
	else
		set the comparison_string to "abcdefghijklmnopqrstuvwxyz"
		set the source_string to "ABCDEFGHIJKLMNOPQRSTUVWXYZ"
	end if
	set the new_text to ""
	repeat with this_char in this_text
		set x to the offset of this_char in the comparison_string
		if x is not 0 then
			set the new_text to (the new_text & character x of the source_string) as string
		else
			set the new_text to (the new_text & this_char) as string
		end if
	end repeat
	return the new_text
end change_case

on replace_chars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replace_chars


(*
The function format_date takes a date object and a format string.
Certain place holders in the format string are replaced by the
appropriate date values. The following placeholders are defined:

<<Year>>: The number of the year (4 digits)
<<Year 2>>: The number of the year (2 digits)
<<Century number>>: Year div 100
<<Century>>: Century number + 1

<<Month>>: Number of month, 1 or 2 digits
<<Month 2>>: Number of month, 2 digits
<<Month name>>: Name of the month. Localized, if locale files are installed, else English

<<Day>>: Day of month, 1 or 2 digits
<<Day 2>>: Day of month, 2 digits
<<Day name>>: Name of the weekday. Localized, if locale files are installed, else English
<<Day suffix>>: st, nd, rd or th depending on last digit of day number
	(Note: This is not localized since such suffixes have several special
	grammatical aspects in other languages particularly if there is gender and case)

<<Hour>>: Hour of the day, 1 or 2 digits
<<Hour 2>>: Hour of the day, 2 digits

<<Minute>>: Minute of the hour, 1 or 2 digits
<<Minute 2>>: Minute of the hour, 2 digits

<<Second>>: Second of the minute, 1 or 2 digits
<<Second 2>>: Second of the minute, 2 digits

Installation of locale files for script bundles and apps is described at
http://www.j-schell.de/node/106 (English)
http://www.j-schell.de/node/105 (German)

If this file is the original bundly I provided,
open the ÒBundle ContentsÓ to see the locale structure
for English, French, German and Dutch.

© 2010, JŸrgen Schell
*)



on format_date(the_date, format_string)
	-- store all elements of the date into variables
	set {year_num, month_num, day_num, hour_num, minute_num, second_num, time_string} to Â
		{year of the_date, (month of the_date) as integer, day of the_date, hours of the_date, Â
			minutes of the_date, seconds of the_date, time string of the_date}
	set {month_name, day_name} to {(month of the_date) as string, (weekday of the_date) as string}
	set suffix_list to {"th", "st", "nd", "rd", "th", "th", "th", "th", "th", "th"}
	
	set placeholders to Â
		{"<<Year>>", Â
			"<<Year 2>>", Â
			"<<Century number>>", Â
			"<<Century>>", Â
			"<<Month>>", Â
			"<<Month 2>>", Â
			"<<Month name>>", Â
			"<<Day>>", Â
			"<<Day 2>>", Â
			"<<Day name>>", Â
			"<<Day suffix>>", Â
			"<<Hour>>", Â
			"<<Hour 2>>", Â
			"<<Minute>>", Â
			"<<Minute 2>>", Â
			"<<Second>>", Â
			"<<Second 2>>", Â
			"<<Time>>"}
	
	set value_list to {Â
		year_num as string, Â
		dd(year_num), Â
		(year_num div 100) as string, Â
		(year_num div 100 + 1) as string, Â
		month_num as string, Â
		dd(month_num), Â
		localized string of (month_name), Â
		day_num as string, Â
		dd(day_num), Â
		localized string of (day_name), Â
		item (day_num mod 10 + 1) of suffix_list, Â
		hour_num as string, Â
		dd(hour_num), Â
		minute_num as string, Â
		dd(minute_num), Â
		second_num as string, Â
		dd(second_num), Â
		time_string as string}
	
	-- repace elements of format string
	set old_delims to AppleScript's text item delimiters
	-- in a loop, replace all placeholders by the
	-- values from value_list
	repeat with x from 1 to count placeholders
		set ph to item x of placeholders
		set val to item x of value_list
		set AppleScript's text item delimiters to ph
		set temp to every text item of format_string
		set AppleScript's text item delimiters to val as text
		set format_string to temp as text
	end repeat
	
	set AppleScript's text item delimiters to old_delims
	return format_string
end format_date
on dd(the_num) -- double digits
	set the_val to (text -2 thru -1 of ((the_num + 100) as text)) as text
	return the_val
end dd

--function for cleaning the characters from the file name
on clean_filename(theName)
	--set the list of characters you want to replace
	--disallowedChars will be replaced with the replacementChar
	--in this case, an underscore
	set disallowedChars to ":;,/|!@#$%^&*()+"
	
	--anything in disallowedChars2 will be removed altogether
	set disallowedChars2 to "'"
	
	--set the character you'd like to use to replace the invalid 
	--characters specified in disallowedChars
	set replacementCharacter to "_"
	
	set newName to ""
	
	repeat with i from 1 to length of theName
		
		--check if the character is in disallowedChars
		--replace it with the replacementCharacter if it is
		if ((character i of theName) is in disallowedChars) then
			set newName to newName & replacementCharacter
			
			--check if the character is in disallowedChars2
			--remove it completely if it is
		else if ((character i of theName) is in disallowedChars2) then
			set newName to newName & ""
			
			--if the character is not in either disallowedChars or
			--disallowedChars2, keep it in the file name
		else
			set newName to newName & character i of theName
			
		end if
	end repeat
	
	return newName
end clean_filename







