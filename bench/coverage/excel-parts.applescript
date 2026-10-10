-- Open a book, add notes, hyperlinks and frozen panes through Excel's object model, apply one structure operation,
-- read every sheet's notes, hyperlinks and panes, save the book as xlsx under another name and close it. Alerts are off
-- around the operation; an operation Excel refuses comes back as `error`, not a dialog.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, what to make
-- as JSON ([["note", sheet, cell, text], ["link", sheet, cell, address, sub address, screen tip], ["freeze", sheet,
-- rows, columns]]), the operation as JSON ([op, sheet, range, target sheet or new name, target], or []).
-- Output: JSON {"error": text, "sheets": [[name, [[address, author, text]], [[range, address, sub address, screen tip,
-- text to display]], [freeze panes, split row, split column, the first row and column of the window, the first row and
-- column of its first pane]]]}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on parsed(t)
	set d to (current application's NSString's stringWithString:t)'s dataUsingEncoding:4
	return (current application's NSJSONSerialization's JSONObjectWithData:d options:0 |error|:(missing value)) as list
end parsed

-- "" for a missing value
on txt(v)
	if v is missing value then return ""
	return v as text
end txt

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set makes to parsed(item 4 of argv)
	set opArgs to parsed(item 5 of argv)
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set opError to ""
	set readError to ""
	set sheetsOut to {}
	tell application "Microsoft Excel"
		with timeout of 90 seconds
			set wb to open workbook workbook file name hfs
			set stage to "make"
			try
				repeat with m in makes
					set kind to item 1 of m
					set ws to worksheet (item 2 of m) of wb
					if kind is "note" then
						add comment (range (item 3 of m) of ws) comment text (item 4 of m)
					else if kind is "link" then
						set props to {address:(item 4 of m), sub address:(item 5 of m)}
						if (item 6 of m) is not "" then set props to props & {screen tip:(item 6 of m)}
						make new hyperlink at (range (item 3 of m) of ws) with properties props
					else if kind is "freeze" then
						activate object ws
						set w to active window
						set freeze panes of w to false
						set split row of w to (item 3 of m)
						set split column of w to (item 4 of m)
						set freeze panes of w to true
					end if
				end repeat
				if (count of opArgs) > 0 then
					set stage to "op"
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
			set stage to "sheets"
			try
				repeat with i from 1 to (count of worksheets of wb)
					set w to worksheet i of wb
					set stage to "notes"
					set notesOut to {}
					set addrText to ""
					-- no cell with a note: Excel answers a parameter error
					try
						set addrText to get address (special cells (used range of w) type cell type comments)
					end try
					if addrText is not "" then
						set AppleScript's text item delimiters to ","
						set areaList to text items of addrText
						set AppleScript's text item delimiters to ""
						repeat with ad in areaList
							set ar to range (ad as text) of w
							repeat with j from 1 to (count of cells of ar)
								set ca to get address (cell j of ar)
								set cm to Excel comment of (range ca of w)
								set end of notesOut to {ca, my txt(author of cm), my txt(Excel comment text cm)}
							end repeat
						end repeat
					end if
					set stage to "links"
					set linksOut to {}
					repeat with k from 1 to (count of hyperlinks of w)
						set h to hyperlink k of w
						set r to range object of h
						set end of linksOut to {(get address r), my txt(address of h), my txt(sub address of h), my txt(screen tip of h), my txt(text to display of h)}
					end repeat
					set stage to "panes"
					activate object w
					set aw to active window
					set pn to pane 1 of aw
					set panesOut to {freeze panes of aw, split row of aw, split column of aw, scroll row of aw, scroll column of aw, scroll row of pn, scroll column of pn}
					set end of sheetsOut to {name of w, notesOut, linksOut, panesOut}
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
	set d0 to current application's NSDictionary's dictionaryWithObjects:{opError, sheetsOut} forKeys:{"error", "sheets"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
