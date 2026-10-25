-- Open a book Excel made for the keep step, apply one structure operation as `excel-struct.applescript` does, save it
-- as xlsx under another name and close it. Alerts are off around the operation (deleting a sheet asks otherwise); an
-- operation Excel refuses answers its error. Arguments: the book's path and the path to save to (inside Excel's
-- sandbox), the saved file's name, the op, its sheet, its range (whole rows and columns as cell ranges), the target
-- sheet or the new name, the target cell.
on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set op to item 4 of argv
	set argSheet to item 5 of argv
	set argRange to item 6 of argv
	set argTarget to item 7 of argv
	set argTo to item 8 of argv
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set opError to ""
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			set wb to open workbook workbook file name hfs
			set display alerts to false
			try
				set ws to worksheet argSheet of wb
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
				else if op is "renameSheet" then
					set name of ws to argTarget
				else if op is "deleteSheet" then
					delete ws
				else
					error "unknown op " & op
				end if
			on error msg number k
				set opError to (k as text) & " " & msg
			end try
			set display alerts to true
			calculate full
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	if opError is "" then return "done"
	return "error " & opError
end run
