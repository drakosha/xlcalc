-- Open a book, apply one structure operation, compute the book, read what the listed cells show by their conditional
-- formats (the display format: Range.Text, the fill and font colours, bold and italic), save the book as xlsx under
-- another name and close it. Alerts are off around the operation; an operation Excel refuses comes back as `error`.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the
-- operation as JSON ([op, sheet, range, target sheet or new name, target], or []), the cells to read as JSON ([[sheet,
-- range]]).
-- Output: JSON {"error": text, "looks": [[sheet, [[address, text, [r, g, b] of the fill, its colour index, [r, g, b] of
-- the font, bold, italic]]]]}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on parsed(t)
	set d to (current application's NSString's stringWithString:t)'s dataUsingEncoding:4
	return (current application's NSJSONSerialization's JSONObjectWithData:d options:0 |error|:(missing value)) as list
end parsed

on txt(v)
	if v is missing value then return ""
	return v as text
end txt

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set opArgs to parsed(item 4 of argv)
	set reads to parsed(item 5 of argv)
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set opError to ""
	set readError to ""
	set looksOut to {}
	tell application "Microsoft Excel"
		with timeout of 120 seconds
			set wb to open workbook workbook file name hfs
			set stage to "op"
			try
				if (count of opArgs) > 0 then
					set op to item 1 of opArgs
					set argSheet to item 2 of opArgs
					set argRange to item 3 of opArgs
					set argTarget to item 4 of opArgs
					set argTo to item 5 of opArgs
					set display alerts to false
					if argSheet is not "" then set ws to worksheet argSheet of wb
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
					set display alerts to true
				end if
			on error msg number k
				set display alerts to true
				set opError to stage & ": " & (k as text) & " " & msg
			end try
			calculate full
			set stage to "looks"
			try
				repeat with readItem in reads
					set ws to worksheet (item 1 of readItem) of wb
					set rng to range (item 2 of readItem) of ws
					set cellsOut to {}
					repeat with j from 1 to (count of cells of rng)
						set oneCell to cell j of rng
						set dispFmt to display format of oneCell
						set intObj to interior object of dispFmt
						set fontObj to font object of dispFmt
						set end of cellsOut to {(get address oneCell), my txt(string value of oneCell), (color of intObj), ((color index of intObj) as text), (color of fontObj), (bold of fontObj), (italic of fontObj)}
					end repeat
					set end of looksOut to {name of ws, cellsOut}
				end repeat
			on error msg number k
				set readError to stage & ": " & (k as text) & " " & msg
			end try
			if readError is "" then
				activate object (worksheet 1 of wb)
				save workbook as wb filename out file format Excel XML file format
				close workbook savedName saving no
			else
				close wb saving no
			end if
		end timeout
	end tell
	if readError is not "" then error "reading: " & readError
	set d0 to current application's NSDictionary's dictionaryWithObjects:{opError, looksOut} forKeys:{"error", "looks"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
