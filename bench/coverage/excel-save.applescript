-- Open a book, compute it fully, save it as xlsx under another name and close it.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name:
-- after `save workbook as` the book is known by that name; and formulas to enter first, as Excel 365 takes a typed
-- one (`formula2`), lines of a range of the first sheet and its formula joined by a tab (none when left out); and
-- cells of the first sheet, by commas, whose formula as Excel 365 shows it (`formula2`) is read after the book is
-- computed: the output is then their lines, each cell and its formula joined by a tab.
on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			-- a book with links to other workbooks opens on the values it stores, as when its prompt is answered Don't Update
			set wb to open workbook workbook file name hfs update links do not update links
			if (count of argv) > 3 then
				repeat with ln in paragraphs of (item 4 of argv)
					if (length of ln) > 0 then
						set AppleScript's text item delimiters to (character id 9)
						set parts to text items of ln
						set AppleScript's text item delimiters to ""
						set formula2 of range (item 1 of parts) of worksheet 1 of wb to (item 2 of parts)
					end if
				end repeat
			end if
			calculate full
			set shown to {}
			if (count of argv) > 4 then
				set AppleScript's text item delimiters to ","
				set cells to text items of (item 5 of argv)
				set AppleScript's text item delimiters to ""
				repeat with a in cells
					set end of shown to (a as text) & (character id 9) & (formula2 of range (a as text) of worksheet 1 of wb)
				end repeat
			end if
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	if shown is {} then return "saved"
	set AppleScript's text item delimiters to (character id 10)
	set t to shown as text
	set AppleScript's text item delimiters to ""
	return t
end run
