-- choosing folder to set up issue, volume, and year
-- my_folder is also used as location for files
choose folder with prompt "Choose a folder" default location (path to desktop as alias)
set my_folder to result
tell application "Finder" to set folder_name to the name of (folder my_folder)
if (count (every character in folder_name)) > 5 then
	set issue_number to ((text 3 thru 5 of folder_name) as number) Â
		& "-" & Â
		((text 7 thru -1 of folder_name) as number) as string
else
	set issue_number to (((text 3 thru 5 of folder_name) as number) as string)
end if
copy ((text ((offset of "-" in issue_number) + 1) thru -1 of issue_number) as number) to issue_num
if (issue_num mod 4) = 0 then
	set volume_number to (issue_num / 4) as integer
else
	set volume_number to ((issue_num div 4) + 1) as integer
end if
set volume_year to 1957 + volume_number
set begin_xml to "<issue><volume>" & (volume_number as string) & "</volume><number>" & (issue_number as string) & "</number><year>" & (volume_year as string) & "</year><open_access />"
display dialog "setting issue-id as:" default answer begin_xml
set begin_xml to text returned of the result
-- start with writing to file
set file_path to my_folder & issue_number & ".xml" as string
set file_ref to open for access file_path with write permission
write begin_xml to file_ref
close access file_ref

-- choosing file in chosen location. Will be replaced by repeat loop
-- choose file with prompt "Choose file" default location my_folder as alias -- legacy, now replaced with loop
tell application "Finder"
	set current_batch to (document files of entire contents of my_folder whose name ends with "pdf")
	repeat with i from 1 to count (every item of current_batch)
		set my_file to item i of current_batch
		
		-- actual collection of content of pdf. All wrapped in tell to Skim
		tell application "Skim"
			activate
			open my_file as alias
			-- add stuff to set size of window
			tell application "System Events"
				keystroke "0" using command down
			end tell
			tell document 1
				get paragraph 1 of page 1
				set this_title to result as string
				get paragraph -1 of page 1
				set first_page to result
				get properties
				get info of result
				set no_pages to page count of result
			end tell
			display dialog "Confirm title" buttons "OK" default button "OK" default answer this_title
			set this_title to text returned of the result
			display dialog "Confirm first page" buttons "OK" default button "OK" default answer first_page
			set first_page to text returned of the result as integer
			set the clipboard to ""
			set old_clipboard to the clipboard
			display alert "copy author name" buttons "OK" default button "OK" giving up after 3
			repeat while old_clipboard = (the clipboard)
				delay 1
			end repeat
			set this_author to the clipboard
			set first_name to first word of this_author
			set name_length to count words in this_author
			set middle_name to ""
			if name_length = 3 then
				set middle_name to second word of this_author
			end if
			set last_name to last word of this_author
			display dialog "Confirm first name" buttons "OK" default button "OK" default answer first_name
			set first_name to text returned of the result
			display dialog "Confirm middle name" buttons "OK" default button "OK" default answer middle_name
			set middle_name to text returned of the result
			display dialog "Confirm last name" buttons "OK" default button "OK" default answer last_name
			set last_name to text returned of the result
			set the clipboard to ""
			set old_clipboard to the clipboard
			display dialog "Copy additional info" buttons "OK" default button "OK"
			repeat while old_clipboard = (the clipboard)
				delay 1
			end repeat
			set add_info to the clipboard
			display dialog "Confirm affiliation/country" buttons {"Empty", "Affiliation", "Country"} default answer add_info -- to be checked what happens if empty
			copy the result as list to {add_info, info_type}
			if info_type = "Affiliation" then
				set info_tag to "<affiliation>"
				set cl_info_tag to "</affiliation>"
			else
				set info_tag to "<country>"
				set cl_info_tag to "</country>"
			end if
			set this_xml to (("<title>" & this_title & "</title>" & return & "<author> <first name> " & first_name & " </first name><middle name> " & middle_name & "</middle name><last name>" & last_name & "</last name></author>" & return & Â
				info_tag & add_info & cl_info_tag & return & Â
				"<pages>" & first_page as string) & "-" & first_page + (no_pages - 1) as string) & "</pages>"
			display dialog "Confirm XML" buttons "OK" default button "OK" default answer this_xml
			close document 1
		end tell
		-- writing result to file
		set file_ref to open for access file_path with write permission
		write this_xml to file_ref starting at eof
		close access file_ref
		
	end repeat
end tell
-- write closing tags
set file_ref to open for access file_path with write permission
write "</issue>" to file_ref starting at eof
close access file_ref

