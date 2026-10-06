-- Open a book, apply one structure operation, read the formulas of a rectangle of every sheet, the extra
-- cells and the defined names, close without saving. Alerts are off around the operation (deleting a
-- sheet asks for confirmation otherwise); an operation Excel refuses comes back as `error`, not a dialog.
-- Arguments: book path (inside Excel's sandbox), op, sheet, range (cells of the rows or columns to insert
-- or delete; whole rows and columns written as cell ranges), target sheet or new name, target range, the rectangle to read on every sheet, extra cells as lines "sheet<TAB>address".
-- Output: JSON {"error": text, "sheets": [[name, used range, formulas of the rectangle]],
-- "extra": [[sheet, address, formula]], "names": [[name, references]]}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set op to item 2 of argv
	set argSheet to item 3 of argv
	set argRange to item 4 of argv
	set argTarget to item 5 of argv
	set argTo to item 6 of argv
	set rect to item 7 of argv
	set extraText to item 8 of argv
	set hfs to (POSIX file p) as text
	set opError to ""
	set readError to ""
	set sheetsOut to {}
	set extraOut to {}
	set namesOut to {}
	tell application "Microsoft Excel"
		with timeout of 60 seconds
			set wb to open workbook workbook file name hfs
			set display alerts to false
			set stage to "sheet"
			try
				set ws to worksheet argSheet of wb
				set stage to "op"
				if op is "insertRows" then
					insert into range (entire row of (range argRange of ws)) shift shift down
				else if op is "deleteRows" then
					delete range (entire row of (range argRange of ws)) shift shift up
				else if op is "insertColumns" then
					insert into range (entire column of (range argRange of ws)) shift shift to right
				else if op is "deleteColumns" then
					delete range (entire column of (range argRange of ws)) shift shift to left
				else if op is "moveRange" then
					cut range (range argRange of ws) destination of cut (range argTo of worksheet argTarget of wb)
				else if op is "copyRange" then
					copy range (range argRange of ws) destination (range argTo of worksheet argTarget of wb)
				else if op is "renameSheet" then
					set name of ws to argTarget
				else if op is "deleteSheet" then
					delete ws
				else
					error "unknown op " & op
				end if
			on error msg number k
				set opError to stage & ": " & (k as text) & " " & msg
			end try
			set display alerts to true
			set stage to "sheets"
			try
				repeat with i from 1 to (count of worksheets of wb)
					set w to worksheet i of wb
					set stage to "used range"
					set ur to used range of w
					set u to get address ur
					set stage to "formulas"
					set f to formula of (range rect of w)
					set end of sheetsOut to {name of w, u, f}
				end repeat
				set stage to "extra"
				repeat with ln in paragraphs of extraText
					if (length of ln) > 0 then
						set AppleScript's text item delimiters to (character id 9)
						set parts to text items of ln
						set AppleScript's text item delimiters to ""
						set fx to ""
						try
							set fx to formula of range (item 2 of parts) of worksheet (item 1 of parts) of wb
						end try
						set end of extraOut to {item 1 of parts, item 2 of parts, fx}
					end if
				end repeat
				set stage to "names"
				repeat with i from 1 to (count of named items of wb)
					set nm to named item i of wb
					set end of namesOut to {name of nm, references of nm}
				end repeat
			on error msg number k
				set readError to stage & ": " & (k as text) & " " & msg
			end try
			close wb saving no
		end timeout
	end tell
	if readError is not "" then error "reading: " & readError
	set d0 to current application's NSDictionary's dictionaryWithObjects:{opError, sheetsOut, extraOut, namesOut} forKeys:{"error", "sheets", "extra", "names"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
