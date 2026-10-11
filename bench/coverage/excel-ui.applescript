-- Open a ui case's book (sheet "S"), compute it fully and take the case's action with Excel in front, on the keyboard
-- (System Events, the handlers of excel-keys.applescript): "bars" selects each cell in turn and reads the formula bar,
-- then presses F2 and reads the bar and the element in focus, and Escape; "paste" puts the text on the clipboard as plain
-- text and presses Cmd+V (the V key: any layout) with the cell selected; "fill" selects the range, types the text and presses Control+Return.
-- Windows Excel shows for a paste or a fill are answered (`answerDialog`). Then read each cell: what it shows (`string
-- value`), its format codes, prefix character, value (a number only, not read from a cell that shows an error), formula
-- (`formula`, `formula local`, `formula2`, `formula2 local`) and whether it is part of a formula of the old kind (`has
-- array`, with its `formula array`); each column's `column width`, `width` and `use standard width`; save the book as
-- xlsx under another name and close it.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the action as
-- JSON {"do": "bars" | "paste" | "fill" | "none", "cells": [...], "at": cell, "range": range, "text": text}, the cells to
-- read and the columns as JSON arrays, the path of excel-keys.applescript. "none" takes no action: the book is only read.
-- Output: JSON {"dialog": text, "bars": [[[description, value] selected, [description, value] in editing, focus, dialog],
-- ...], "cells": [[text, code, local, prefix, value, formula, formula local, formula2, formula2 local, has array, formula
-- array], ...], "columns": [[column width, points, standard], ...]}.
use AppleScript version "2.4"
use framework "Foundation"
use framework "AppKit"
use scripting additions

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set act to my fromJSON(item 4 of argv)
	set reads to my fromJSON(item 5 of argv) as list
	set cols to my fromJSON(item 6 of argv) as list
	set kb to my keysLib(item 7 of argv)
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set todo to (act's objectForKey:"do") as text
	set said to ""
	set barsOut to {}
	set cellsOut to {}
	set colsOut to {}
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			set wb to open workbook workbook file name hfs
			set ws to worksheet "S" of wb
			set bookName to name of wb
			calculate full
		end timeout
	end tell
	try
		if todo is "bars" then
			set barsOut to my bars(ws, (act's objectForKey:"cells") as list, bookName, kb)
		else if todo is "paste" then
			my toClipboard((act's objectForKey:"text") as text)
			tell application "Microsoft Excel" to select range ((act's objectForKey:"at") as text) of ws
			kb's toFront()
			-- the V key, whatever the layout types on it, as a person pastes
			tell application "System Events" to key code 9 using command down
			delay 1.5
			set said to kb's answerDialog()
		else if todo is "fill" then
			tell application "Microsoft Excel" to select range ((act's objectForKey:"range") as text) of ws
			set said to kb's typeFill((act's objectForKey:"text") as text)
		end if
	on error msg number k
		-- a step that failed (Excel's window on another desktop) leaves no book open
		tell application "Microsoft Excel" to close wb saving no
		error msg number k
	end try
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			calculate
			repeat with a in reads
				set rg to range (a as text) of ws
				set tx to string value of rg
				if tx is missing value then set tx to ""
				set pc to prefix character of rg
				if pc is missing value then set pc to ""
				-- an error's value is not read (a range's errors crashed Excel); a number too wide for the column shows ## and is read
				set vv to ""
				if tx does not start with "#" or tx starts with "##" then
					try
						set vv to value of rg
					end try
					-- a number alone: a date comes as a date object, which JSON has no place for
					if class of vv is not real and class of vv is not integer then set vv to ""
				end if
				set ha to has array of rg
				set fa to ""
				if ha then set fa to formula array of rg
				set end of cellsOut to {tx as text, number format of rg, number format local of rg, pc as text, vv, formula of rg, formula local of rg, formula2 of rg, formula2 local of rg, ha, fa}
			end repeat
			repeat with c in cols
				set firstCell to range ((c as text) & "1") of ws
				set end of colsOut to {column width of firstCell, width of firstCell, use standard width of firstCell}
			end repeat
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{said, barsOut, cellsOut, colsOut} forKeys:{"dialog", "bars", "cells", "columns"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run

-- each cell of `cells` selected in turn with Excel in front: the formula bar then, and once F2 began the cell's editing,
-- with the element in focus; Escape ends the editing, and a window Excel shows meanwhile is answered
on bars(ws, cells, bookName, kb)
	set bar to kb's formulaBar(bookName)
	set out to {}
	repeat with a in cells
		tell application "Microsoft Excel" to select range (a as text) of ws
		set onSelect to kb's barText(bar)
		kb's toFront()
		tell application "System Events" to key code 120
		delay 0.5
		set onEdit to kb's barText(bar)
		set inFocus to kb's focusedText()
		tell application "System Events" to key code 53
		delay 0.5
		set end of out to {onSelect, onEdit, inFocus, kb's answerDialog()}
	end repeat
	return out
end bars

-- the text on the general pasteboard as plain text, nothing else with it
on toClipboard(t)
	set pb to current application's NSPasteboard's generalPasteboard()
	pb's clearContents()
	pb's setString:t forType:(current application's NSPasteboardTypeString)
end toClipboard

on fromJSON(s)
	set js to current application's NSString's stringWithString:s
	return current application's NSJSONSerialization's JSONObjectWithData:(js's dataUsingEncoding:4) options:0 |error|:(missing value)
end fromJSON

-- the handlers of excel-keys.applescript (its path `p`) as a script object
on keysLib(p)
	set src to (current application's NSString's stringWithContentsOfFile:p encoding:4 |error|:(missing value)) as text
	return run script ("script K" & linefeed & src & linefeed & "end script" & linefeed & "return K")
end keysLib
