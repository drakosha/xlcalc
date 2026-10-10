-- Open a book, note its active sheet, set the visibility of sheets through Excel's object model, apply one sheet
-- operation, recalculate, read every sheet's visibility, the active sheet and the values of the read cells, save the book
-- as xlsx under another name and close it. Alerts are off around the steps; a step Excel refuses comes back as `error`,
-- not a dialog, and the steps after it are not taken.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the
-- visibilities to set as JSON ([[sheet, -1 visible | 0 hidden | 2 very hidden]]), the operation as JSON ([op, sheet,
-- range, target sheet or new name, target], or []), the read sheet's name and its cells as JSON (["A1", ...]).
-- Output: JSON {"error": text, "opened": name, "active": name, "sheets": [[name, visibility]], "values": [value]}; a
-- value is a number, a text, a boolean or "" for a blank cell (the read cells hold no errors).
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on parsed(t)
	set d to (current application's NSString's stringWithString:t)'s dataUsingEncoding:4
	return (current application's NSJSONSerialization's JSONObjectWithData:d options:0 |error|:(missing value)) as list
end parsed

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set makes to parsed(item 4 of argv)
	set opArgs to parsed(item 5 of argv)
	set readSheet to item 6 of argv
	set readCells to parsed(item 7 of argv)
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set stepError to ""
	set sheetsOut to {}
	set valuesOut to {}
	tell application "Microsoft Excel"
		with timeout of 90 seconds
			set wb to open workbook workbook file name hfs
			set openedName to name of active sheet of wb
			set stage to "make"
			set display alerts to false
			try
				repeat with m in makes
					set ws to worksheet (item 1 of m) of wb
					set vis to (item 2 of m) as integer
					if vis is -1 then
						set visible of ws to sheet visible
					else if vis is 0 then
						set visible of ws to sheet hidden
					else
						set visible of ws to sheet very hidden
					end if
				end repeat
				if (count of opArgs) > 0 then
					set stage to "op"
					set op to item 1 of opArgs
					set argSheet to item 2 of opArgs
					set argTarget to item 4 of opArgs
					if op is "renameSheet" then
						set name of worksheet argSheet of wb to argTarget
					else if op is "deleteSheet" then
						delete worksheet argSheet of wb
					else if op is "insertSheet" then
						if argSheet is "" then
							set nw to make new worksheet at after (worksheet (count of worksheets of wb) of wb)
						else
							set nw to make new worksheet at before (worksheet argSheet of wb)
						end if
						set name of nw to argTarget
					else if op is "moveSheet" then
						if argTarget is "" then
							move worksheet argSheet of wb to after (worksheet (count of worksheets of wb) of wb)
						else
							move worksheet argSheet of wb to before (worksheet argTarget of wb)
						end if
					else
						error "unknown op " & op
					end if
				end if
			on error msg number k
				set stepError to stage & ": " & (k as text) & " " & msg
			end try
			set display alerts to true
			calculate
			repeat with i from 1 to (count of worksheets of wb)
				set w to worksheet i of wb
				set v to visible of w
				if v is sheet visible then
					set vn to -1
				else if v is sheet hidden then
					set vn to 0
				else
					set vn to 2
				end if
				set end of sheetsOut to {name of w, vn}
			end repeat
			set activeName to name of active sheet of wb
			repeat with ca in readCells
				set cv to value of range (ca as text) of worksheet readSheet of wb
				if cv is missing value then set cv to ""
				set end of valuesOut to cv
			end repeat
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{stepError, openedName, activeName, sheetsOut, valuesOut} forKeys:{"error", "opened", "active", "sheets", "values"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
