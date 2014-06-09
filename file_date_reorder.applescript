rename_file()

on rename_file()
	tell application "Finder"
		set files_ to choose file with multiple selections allowed
		
		repeat with file_ in files_ 
		set file_alias to file_ as alias
		set display_name to name of file_alias
		set org_date to (characters 1 thru 10 of display_name) as string
		log (display_name)
		
		set month_ to (characters 1 thru 2 of display_name) as string
		log (month_)
		set day_ to (characters 4 thru 5 of display_name) as string
		log (day_)
		set year_ to (characters 7 thru 10 of display_name) as string
		log (year_)
		
		set new_date_ to (year_ & "-" & month_ & "-" & day_) as string
		try
			set display_name to my replace_chars(display_name, org_date, new_date_)
		on error error_message
			log (error_message)
		end try
		
		log (display_name)
		set name of file_ to display_name
		
		end repeat 
		
		
	end tell
	
	
end rename_file

on replace_chars(this_text, search_string, replacement_string)
	set AppleScript's text item delimiters to the search_string
	set the item_list to every text item of this_text
	set AppleScript's text item delimiters to the replacement_string
	set this_text to the item_list as string
	set AppleScript's text item delimiters to ""
	return this_text
end replace_chars