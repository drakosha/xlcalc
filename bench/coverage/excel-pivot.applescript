-- Open a book Excel for Windows saved with pivot tables, read what each listed cell shows as it opened (no
-- recalculation, no refresh), save it as xlsx under another name and close it. Arguments: the book's path and the path
-- to save to (inside Excel's sandbox), the saved file's name, the cells as "P!A3;P!B3;G!A1". Output: JSON
-- [[text, value], ...] in their order, the value as text ("" for an empty cell).
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set AppleScript's text item delimiters to ";"
	set cellList to text items of (item 4 of argv)
	set AppleScript's text item delimiters to ""
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set got to {}
	tell application "Microsoft Excel"
		with timeout of 170 seconds
			-- links to other books stay as saved: Excel asks about updating them otherwise
			set wb to open workbook workbook file name hfs update links do not update links
			repeat with c in cellList
				set AppleScript's text item delimiters to "!"
				set sh to text item 1 of (c as text)
				set ad to text item 2 of (c as text)
				set AppleScript's text item delimiters to ""
				set rg to range ad of worksheet sh of wb
				set tx to string value of rg
				if tx is missing value then set tx to ""
				set v to value of rg
				if v is missing value then set v to ""
				set end of got to {tx as text, v as text}
			end repeat
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:got options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
