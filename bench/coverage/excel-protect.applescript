-- Open a book, take the steps (a copy, with Paste Special or Range.Copy to a destination; rows inserted; a range's formats
-- cleared; a range's locked and formula hidden set; a sheet protected as the Protect Sheet dialog does, with what it
-- allows, or unprotected), read every sheet's protection and the listed cells' locked and formula hidden, save the book
-- as xlsx under another name and close it. A step Excel refuses comes back as `error`.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the steps as
-- JSON ([["copy", sheet, range, target sheet, target range, "all", "formats", "values" or "" for Range.Copy],
-- ["insertRows", sheet, rows], ["clearFormats", sheet, range], ["locked", sheet, range, locked, formula hidden],
-- ["protect", sheet, [the allowed: "formatCells", "insertRows", …]], ["unprotect", sheet]]), the cells to read as JSON
-- ([[sheet, range]]).
-- Output: JSON {"error": text, "sheets": [[name, contents, drawing objects, scenarios, enable selection,
-- [formatting cells, formatting columns, formatting rows, inserting columns, inserting rows, inserting hyperlinks,
-- deleting columns, deleting rows, sorting, filtering, using pivot table]]], "cells": [[sheet, [[address, locked,
-- formula hidden]]]]}.
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
		if t is "formats" then return paste formats
		return paste all
	end tell
end pasteType

-- whether a list of texts holds a text
on allows(names, n)
	repeat with x in names
		if (x as text) is n then return true
	end repeat
	return false
end allows

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set steps to parsed(item 4 of argv)
	set reads to parsed(item 5 of argv)
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set stepError to ""
	set readError to ""
	set sheetsOut to {}
	set cellsOut to {}
	tell application "Microsoft Excel"
		with timeout of 120 seconds
			set wb to open workbook workbook file name hfs
			repeat with stepItem in steps
				set what to (item 1 of stepItem) as text
				try
					set display alerts to false
					set ws to worksheet ((item 2 of stepItem) as text) of wb
					activate object ws
					if what is "copy" then
						set wt to worksheet ((item 4 of stepItem) as text) of wb
						if (item 6 of stepItem) as text is "" then
							copy range (range ((item 3 of stepItem) as text) of ws) destination (range ((item 5 of stepItem) as text) of wt)
						else
							copy range (range ((item 3 of stepItem) as text) of ws)
							activate object wt
							paste special (range ((item 5 of stepItem) as text) of wt) what (my pasteType((item 6 of stepItem) as text)) operation paste special operation none skip blanks false transpose false
						end if
						set cut copy mode to false
					else if what is "insertRows" then
						insert into range (entire row of (range ((item 3 of stepItem) as text) of ws)) shift shift down
					else if what is "clearFormats" then
						clear range formats (range ((item 3 of stepItem) as text) of ws)
					else if what is "locked" then
						set rg to range ((item 3 of stepItem) as text) of ws
						set locked of rg to (item 4 of stepItem)
						set formula hidden of rg to (item 5 of stepItem)
					else if what is "protect" then
						set ok to item 3 of stepItem
						protect worksheet ws drawing objects true worksheet contents true scenarios true allow formatting cells (my allows(ok, "formatCells")) allow formatting columns (my allows(ok, "formatColumns")) allow formatting rows (my allows(ok, "formatRows")) allow inserting columns (my allows(ok, "insertColumns")) allow inserting rows (my allows(ok, "insertRows")) allow inserting hyperlinks (my allows(ok, "insertHyperlinks")) allow deleting columns (my allows(ok, "deleteColumns")) allow deleting rows (my allows(ok, "deleteRows")) allow sorting (my allows(ok, "sort")) allow filtering (my allows(ok, "autoFilter")) allow using pivot table (my allows(ok, "pivotTables"))
					else if what is "unprotect" then
						unprotect ws
					end if
					set display alerts to true
				on error msg number k
					set display alerts to true
					set stepError to stepError & what & " " & ((item 2 of stepItem) as text) & ": " & (k as text) & " " & msg & "; "
				end try
			end repeat
			try
				-- `repeat with ws in (worksheets of wb)` gives a parameter error
				repeat with i from 1 to (count of worksheets of wb)
					set ws to worksheet i of wb
					set po to protection object of ws
					set allowOut to {allow formatting cells of po, allow formatting columns of po, allow formatting rows of po, allow inserting columns of po, allow inserting rows of po, allow inserting hyperlinks of po, allow deleting columns of po, allow deleting rows of po, allow sorting of po, allow filtering of po, allow using pivot table of po}
					set end of sheetsOut to {name of ws, protect contents of ws, protect drawing objects of ws, protect scenarios of ws, (enable selection of ws) as text, allowOut}
				end repeat
				repeat with readItem in reads
					set ws to worksheet ((item 1 of readItem) as text) of wb
					-- `cell j of` a range of a sheet that is not active gives the active sheet's cell
					activate object ws
					set rng to range ((item 2 of readItem) as text) of ws
					set got to {}
					repeat with j from 1 to (count of cells of rng)
						set oneCell to cell j of rng
						set end of got to {(get address oneCell), locked of oneCell, formula hidden of oneCell}
					end repeat
					set end of cellsOut to {name of ws, got}
				end repeat
			on error msg number k
				set readError to (k as text) & " " & msg
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
	set d0 to current application's NSDictionary's dictionaryWithObjects:{stepError, sheetsOut, cellsOut} forKeys:{"error", "sheets", "cells"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
