-- Open a book, copy a range to the clipboard, paste it with Paste Special at a target, then save the book as xlsx under
-- another name. Alerts are off around the paste; a paste Excel refuses comes back as `error`, not a dialog.
-- Arguments: the book's path and the path to save to (inside Excel's sandbox), the saved file's name, the paste as JSON
-- ([source sheet, source range, target sheet, target range, what (all, values, formulas, formats, validation,
-- formulasAndNumberFormats, valuesAndNumberFormats, columnWidths, allExceptBorders, notes; copy: a plain copy to the
-- target, Range.Copy with a destination), transpose "1" or "0", the operation (none, add, subtract, multiply, divide),
-- skip blanks "1" or "0"]).
-- Output: JSON {"error": text}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on parsed(t)
	set d to (current application's NSString's stringWithString:t)'s dataUsingEncoding:4
	return (current application's NSJSONSerialization's JSONObjectWithData:d options:0 |error|:(missing value)) as list
end parsed

on pasteType(t)
	tell application "Microsoft Excel"
		if t is "values" then return paste values
		if t is "formulas" then return paste formulas
		if t is "formats" then return paste formats
		if t is "validation" then return paste validation
		if t is "formulasAndNumberFormats" then return paste formulas and number formats
		if t is "valuesAndNumberFormats" then return paste values and number formats
		if t is "columnWidths" then return paste column widths
		if t is "allExceptBorders" then return paste all except borders
		if t is "notes" then return paste comments
		if t is "all" then return paste all
	end tell
	error "no paste type " & t
end pasteType

on operationOf(t)
	tell application "Microsoft Excel"
		if t is "add" then return paste special operation add
		if t is "subtract" then return paste special operation subtract
		if t is "multiply" then return paste special operation multiply
		if t is "divide" then return paste special operation divide
		if t is "none" then return paste special operation none
	end tell
	error "no operation " & t
end operationOf

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set op to parsed(item 4 of argv)
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set opError to ""
	tell application "Microsoft Excel"
		with timeout of 120 seconds
			set wb to open workbook workbook file name hfs
			set stage to "copy"
			try
				set display alerts to false
				set ws to worksheet (item 1 of op) of wb
				activate object ws
				set wt to worksheet (item 3 of op) of wb
				if (item 5 of op) is "copy" then
					set stage to "paste"
					copy range (range (item 2 of op) of ws) destination (range (item 4 of op) of wt)
				else
					copy range (range (item 2 of op) of ws)
					set stage to "paste"
					activate object wt
					paste special (range (item 4 of op) of wt) what (my pasteType(item 5 of op)) operation (my operationOf(item 7 of op)) skip blanks ((item 8 of op) is "1") transpose ((item 6 of op) is "1")
				end if
				set display alerts to true
			on error msg number k
				set display alerts to true
				set opError to stage & ": " & (k as text) & " " & msg
			end try
			set cut copy mode to false
			activate object (worksheet 1 of wb)
			select range "A1" of (worksheet 1 of wb)
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{opError} forKeys:{"error"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
