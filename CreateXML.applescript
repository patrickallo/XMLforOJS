-- choosing main folder to start
set main_folder to pick_folder()

-- create issue xml for each folder
tell application "Finder"
repeat with i from 1 to count (every folder of main_folder whose comment is not "processed")
	set issue_folder to folder i of (every folder of main_folder whose comment is not "processed")
	tell me to set issue_number to retrieve_issuenum for (issue_folder as alias)
	tell me to set volume_number to compute_volnum for issue_number
	tell me to set volume_year to compute_year for volume_number
	set begin_xml to "<?xml version=\"1.0\" encoding=\"UTF-8\"?>
	<!DOCTYPE issues PUBLIC \"-//PKP//OJS Articles and Issues XML//EN\" \"http://pkp.sfu.ca/ojs/dtds/2.4/native.dtd\">
	<issues><issue published=\"true\" identification=\"num_vol_year\" current=\"false\"><volume>" & (volume_number as string) & "</volume><number>" & (issue_number as string) & "</number><year>" & (volume_year as string) & "</year><open_access />"
	display dialog "setting issue-id as:" default answer begin_xml
	set begin_xml to text returned of the result
	set end_xml to "</section></issue></issues>"
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
				keystroke "1" using command down -- selecting text tool
			end tell
			set current_pdf to document 1
		end tell
			tell me to handle_section for current_pdf
			tell me to write_to_xml for "<article>"
			tell me to write_title for current_pdf
			tell me to write_author for current_pdf
			tell me to write_add_authors for current_pdf
			tell me to write_pages for current_pdf
			tell me to write_to_xml for "</article>"
			tell application "Skim" to close current_pdf
	end repeat
	
	-- finish xml
	tell me to write_to_xml for end_xml
	-- add comment to folder to keep track of processed folders
	set folder_props to properties of issue_folder
	set comment of folder_props to "processed"
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


-- start/end section and write to file
on handle_section for this_pdf
	global current_section
	copy current_section to previous_section
	set the section_list to {"", "Editorials", "Articles", "Reviews", "Reports", "Announcements"}
	set current_section to (choose from list section_list with prompt ("Section for this article:") default items {previous_section}) as string
	if current_section is previous_section then
	else
		set section_xml to ("</section><section><title>" & current_section & "</title>")
		write_to_xml for section_xml
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
	set title_xml to ("<title>" & this_title & "</title>" & return)
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
	set author_xml to ("<author primary_contact=\"true\"> <first name>" & first item of this_author & "</first name><middle name>" & second item of this_author & "</middle name><last name>" & third item of this_author & "</last name></author>" & return)
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
				set author_xml to ("<author> <first name>" & first item of this_author & "</first name><middle name>" & second item of this_author & "</middle name><last name>" & third item of this_author & "</last name></author>" & return)
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
			set page_range_xml to ("<pages>"&((first_page &"-"&(first_page + no_pages)) as string)&"</pages>")
		end tell
	end tell
	-- display dialog "Confirm pages" buttons "OK" default button "OK" default answer page_range
	-- set page_range to text returned of the result
	write_to_xml for page_range_xml
end write_pages
	
-- write_to_xml
on write_to_xml for this_xml
	global file_path
	set file_ref to open for access file_path with write permission
	write this_xml to file_ref starting at eof
	close access file_ref
end write_to_xml



-- snippets from old version
-- display dialog "Copy additional info" buttons "OK" default button "OK"
-- repeat while old_clipboard = (the clipboard)
-- 	delay 1
-- end repeat
-- set add_info to the clipboard
-- display dialog "Confirm affiliation/country" buttons {"Empty", "Affiliation", "Country"} default answer add_info -- to be checked what happens if empty
-- copy the result as list to {add_info, info_type}
-- if info_type = "Affiliation" then
-- 	set info_tag to "<affiliation>"
-- 	set cl_info_tag to "</affiliation>"
-- else
-- 	set info_tag to "<country>"
-- 	set cl_info_tag to "</country>"
-- end if
-- set this_xml to (("<title>" & this_title & "</title>" & return & "<author> <first name> " & first_name & " </first name><middle name> " & middle_name & "</middle name><last name>" & last_name & "</last name></author>" & return & ¬
-- 	info_tag & add_info & cl_info_tag & return & ¬
-- 	"<pages>" & first_page as string) & "-" & first_page + (no_pages - 1) as string) & "</pages>"
-- display dialog "Confirm XML" buttons "OK" default button "OK" default answer this_xml
