tell application "Finder"
	set target_folder to target of front Finder window
	repeat with i from 1 to count (every folder of target_folder)
		set current_folder to folder i of target_folder
		set current_batch to (document files of entire contents of current_folder whose name ends with "pdf")
		repeat with i from 1 to count (every item of current_batch)
			set current_pdf to item i of current_batch
			tell application "PDFpen"
				activate
				open current_pdf as alias
				tell document 1
					ocr
					repeat while performing ocr
						delay 1
					end repeat
					delay 1
					close with saving
				end tell
			end tell
		end repeat
	end repeat
end tell