-- choosing main folder to start
set main_folder to pick_folder()

-- create issue xml for each folder
tell application "Finder"
repeat with i from 1 to count (every folder of main_folder whose comment is not "processed")
	set issue_folder to folder i of (every folder of main_folder whose comment is not "processed")
	tell me to set issue_number to retrieve_issuenum for (issue_folder as alias)
	global volume_number
	tell me to set volume_number to compute_volnum for issue_number
	global volume_year
	tell me to set volume_year to compute_year for volume_number
	tell me to set pub_date to compute_date(issue_number, volume_number, volume_year)
	set begin_xml to "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" & return & "<!DOCTYPE issues PUBLIC \"-//PKP//OJS Articles and Issues XML//EN\" \"http://pkp.sfu.ca/ojs/dtds/2.4/native.dtd\">" & return & "<issues>" & return & tab & "<issue published=\"true\" identification=\"num_vol_year\" current=\"false\">" & return & tab & tab & "<volume>" & (volume_number as string) & "</volume>" & return & tab & tab & "<number>" & (issue_number as string) & "</number>" & return & tab & tab & "<year>" & (volume_year as string) & "</year>" & return & tab & tab & "<date_published>" & pub_date & "</date_published>" & return & tab & tab & "<open_access/>" & return
	display dialog "setting issue-id as:" default answer begin_xml
	set begin_xml to text returned of the result
	set end_xml to tab & tab & "</section>" & return & tab & "</issue>" & return & "</issues>"
	-- start with writing to file
	global file_path
	set file_path to (issue_folder as alias) & issue_number & ".xml" as string
	set file_ref to open for access file_path with write permission
	write begin_xml to file_ref
	close access file_ref
-- select all pdfs in folder for issue
	set current_batch to (document files of entire contents of issue_folder whose name ends with "pdf")
	global current_section
	set current_section to "" -- will be used by handle_section subroutine
	repeat with i from 1 to count (every item of current_batch)
		set current_file to item i of current_batch
		-- actual collection of content of pdf
		tell application "Skim"
			activate
			open current_file as alias
			set the bounds of the first window to {900, 0, 1920, 1200} -- window to right half of screen
			tell application "System Events"
				keystroke "0" using command down -- pdf to actual size
				delay 1
				keystroke "1" using command down -- selecting text tool
			end tell
			set current_pdf to document 1
		end tell
			tell me to handle_section for current_pdf
			tell me to write_to_xml for tab & tab & tab &"<article locale=\"en_US\" language=\"en\">" & return
			tell me to write_title for current_pdf
			tell me to write_more(current_pdf, "abstract", "Without abstract", "write")
			tell me to write_author for current_pdf
			tell me to write_add_authors for current_pdf
			tell me to write_pages for current_pdf
			tell me to write_to_xml for tab & tab & tab & "</article>" & return
			tell application "Skim" to close current_pdf
	end repeat
	
	-- finish xml
	tell me to write_to_xml for end_xml
	-- add comment to folder to keep track of processed folders -- DOESN'T WORK!!!
	tell issue_folder to set comment to "processed"
	display dialog "Next issue?" buttons {"Yes", "No"} default button "Yes"
	try
		if button returned of the result is "No" then
			exit repeat
		end if
	end try
end repeat
end tell

-----------------------------------------------
-- sub-routines -------------------------------
-----------------------------------------------

-- choosing main folder
on pick_folder()
	choose folder with prompt "Choose a folder" default location (path to desktop as alias)
	return result
end pick_folder

-- get issue_number from foldername
on retrieve_issuenum for a_folder
	tell application "Finder" 
		set folder_name to (name of a_folder)
		if (count (every character in folder_name)) > 5 then
			set issue_number to ((text 3 thru 5 of folder_name) as number) & "-" & ((text 7 thru -1 of folder_name) as number) as string
			else
				set issue_number to (((text 3 thru 5 of folder_name) as number) as string)
		end if
	end tell
	return issue_number
end retrieve_issuenum

-- compute volume_number from issuenumber
on compute_volnum for issue_number
	copy ((text ((offset of "-" in issue_number) + 1) thru -1 of issue_number) as number) to issue_num
		if (issue_num mod 4) = 0 then
			set volume_number to (issue_num / 4) as integer
		else
			set volume_number to ((issue_num div 4) + 1) as integer
		end if
	return volume_number
end compute_volnum

-- compute year from volume
on compute_year for volume_number
	return 1957 + volume_number
end compute_year

-- compute pub_date
on compute_date(an_issue, a_vol, a_year)
	set AppleScript's text item delimiters to "-"
	get first text item of an_issue
	set an_issue to result as number
	set AppleScript's text item delimiters to {""}
	set this_month to ((an_issue div a_vol)*3) 
	return a_year &"-"& this_month &"-1" as string
end compute_date

-- start/end section and write to file
on handle_section for this_pdf
	global current_section
	copy current_section to previous_section
	set the section_list to {"", "Editorials", "Articles", "Reviews", "Reports", "Announcements"}
	set current_section to (choose from list section_list with prompt ("Section for this article:") default items {previous_section}) as string
	if current_section is previous_section then
	else
		if previous_section is "" then -- only open section
			set section_xml to (tab & tab & "<section>" & return & tab & tab & tab & "<title locale=\"en_US\">" & current_section & "</title>" & return & tab & tab & tab & "<abbrev locale=\"en_US\">ED</abbrev>" & return)
			write_to_xml for section_xml
		else -- close section and open section
			set section_xml to (tab & tab & "</section>" & return & tab & tab & "<section>" & return & tab & tab & tab & "<title locale=\"en_US\">" & current_section & "</title>" & return & tab & tab & tab & "<abbrev locale=\"en_US\">ED</abbrev>" & return)
			write_to_xml for section_xml
		end if
	end if
end handle_section
	


-- retrieve and write title
on write_title for this_pdf
	tell application "Skim"
		tell this_pdf
			get paragraph 1 of page 1
			set this_title to result as string
		end tell
	end tell
	display dialog "Confirm title" buttons "OK" default button "OK" default answer this_title
	set this_title to text returned of the result
	-- writing result to file
	set title_xml to (tab & tab & tab & tab & "<title>" & this_title & "</title>" & return)
	write_to_xml for title_xml
end write_title

-- copy and write first author
on write_author for current_pdf
	set the clipboard to ""
	set old_clipboard to the clipboard
	display alert "copy author name" buttons "OK" default button "OK" giving up after 3
	repeat while old_clipboard = (the clipboard)
		delay 1
	end repeat
	set this_author to the clipboard
	set this_author to parse_author for this_author
	set email_xml to write_more(current_pdf, "email", "no@email.here", "return")
	set author_xml to (tab & tab & tab & tab & "<author primary_contact=\"true\">"& return & tab & tab & tab & tab & tab &"<firstname>" & first item of this_author & "</firstname>" & return & tab & tab & tab & tab & tab & "<middlename>" & second item of this_author & "</middlename>" & return & tab & tab & tab & tab & tab & "<lastname>" & third item of this_author & "</lastname>" & return & tab & email_xml & tab & tab & tab & tab & "</author>" & return)
	write_to_xml for author_xml
end write_author

-- copy and write remaining authors
on write_add_authors for current_pdf
	repeat
		set the clipboard to ""
		set old_clipboard to the clipboard
		display alert "copy author name" buttons {"OK", "No more authors"} default button "No more authors"
		set decision to button returned of result
		try
			if decision is "No more authors" then
				exit repeat
			else
				repeat while old_clipboard = (the clipboard)
					delay 1
				end repeat
				set this_author to the clipboard
				set this_author to parse_author for this_author
				set email_xml to write_more(current_pdf, "email", "no@email.here", "return")
				set author_xml to (tab & tab & tab & tab & "<author>"& return & tab & tab & tab & tab & tab &"<firstname>" & first item of this_author & "</firstname>" & return & tab & tab & tab & tab & tab & "<middlename>" & second item of this_author & "</middlename>" & return & tab & tab & tab & tab & tab & "<lastname>" & third item of this_author & "</lastname>" & return & tab & email_xml & tab & tab & tab & tab & "</author>" & return)
				write_to_xml for author_xml
			end if
		end try
	end repeat
end write_add_authors

-- parse_author splits up selected text into first middle and last + checks for confirmation and returns as list
on parse_author for an_author
	set first_name to first word of an_author
	set name_length to count words of an_author
	set middle_name to ""
	if name_length = 3 then
		set middle_name to second word of an_author
	end if
	set last_name to last word of an_author
	display dialog "Confirm full name" buttons "OK" default button "OK" default answer (first_name & "+" & middle_name & "+" & last_name)
	set full_name to (text returned of the result)
	set AppleScript's text item delimiters to "+"
	get every text item of full_name
	set full_name to result
	set AppleScript's text item delimiters to {""}
	return full_name
end parse_author


-- retrieve and write pages
on write_pages for this_pdf
	tell application "Skim"
		tell this_pdf
			get paragraph -1 of page 1
			set first_page to result
			display dialog "Confirm first page" buttons "OK" default button "OK" default answer first_page
			set first_page to ((text returned of result) as number)
			get properties
			get info of result
			set no_pages to page count of result
			set page_range_xml to (tab & tab & tab & tab &"<pages>"&((first_page &"-"&(first_page + no_pages)) as string)&"</pages>" & return)
		end tell
	end tell
	-- display dialog "Confirm pages" buttons "OK" default button "OK" default answer page_range
	-- set page_range to text returned of the result
	write_to_xml for page_range_xml
end write_pages

-- retrieve and write optional with default "empty-text" and choice between store and write
on write_more(this_pdf, content_type, default_value, return_or_write)
	global volume_year
	global volume_number
	if ((content_type is "email" and volume_year < 1990) or (content_type is "abstract" and volume_number < 120)) then
		set this_content to default_value
	else
		set the clipboard to ""
		set old_clipboard to the clipboard
		set no_thanks to ("Without " & content_type)
		display alert ("Copy " & content_type & "?") buttons {"OK", no_thanks} default button no_thanks
		set decision to button returned of result
		if decision is no_thanks then
			set this_content to default_value
		else
			repeat while old_clipboard = (the clipboard)
				delay 1
			end repeat
			set this_content to the clipboard
		end if
		display dialog ("Confirm " & content_type) buttons "OK" default button "OK" default answer this_content
		set this_content to (text returned of result)
	end if
	set open_tag to "<" & content_type & ">"
	set close_tag to "</" & content_type & ">"
	set more_content_xml to (tab & tab & tab & tab &open_tag&this_content&close_tag& return)
    if return_or_write is "return" then
		return more_content_xml
	else
		write_to_xml for more_content_xml
	end if
end write_more

	
-- write_to_xml
on write_to_xml for this_xml
	global file_path
	set file_ref to open for access file_path with write permission
	write this_xml to file_ref starting at eof
	close access file_ref
end write_to_xml
