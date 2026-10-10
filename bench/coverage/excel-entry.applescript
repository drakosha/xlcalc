-- Open an entry part's book, type each case's text into its cell (column A of the sheet "S", row by row) as Excel 365
-- takes typed text in the user's language (`formula2 local`), alerts off, then read each cell: what it shows (`string
-- value`), its format codes (`number format`, `number format local`), its prefix character and its value (`value`, a
-- number only, not read from a cell that shows an error); save the book as xlsx under another name and close it.
-- With a fifth argument "keys" the texts are typed on the keyboard instead (System Events, Excel in front, checked
-- before each text; an empty text is not typed), each followed by Return; the window Excel shows for a formula it does
-- not take is answered (`answerDialog`) and its text is the text's refusal. With a last argument "edit" each cell's
-- `formula` and `formula local` are read too, and with Excel in front each cell is selected in turn and its formula bar
-- read (System Events: the description of the window's text area, "Formula Bar. A,1 <text>"), as the object model
-- gives a date's or a percent's number there.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the texts as
-- a JSON array of strings, "keys" and the path of excel-keys.applescript, or nothing; then "edit", or nothing.
-- Output: JSON {"refused": [Excel's error for each text, "" when it took it], "cells": [[text, code, local, prefix, value(,
-- formula, formula local, formula bar)], ...]}, the value "" for a cell whose value is no number (or a date), or that
-- shows an error.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set js to current application's NSString's stringWithString:(item 4 of argv)
	set texts to (current application's NSJSONSerialization's JSONObjectWithData:(js's dataUsingEncoding:4) options:0 |error|:(missing value)) as list
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set refusedOut to {}
	set cellsOut to {}
	set n to count of texts
	set byKeys to (count of argv) > 4 and item 5 of argv is "keys"
	set withEdit to (count of argv) > 4 and item -1 of argv is "edit"
	if byKeys then set kb to my keysLib(item 6 of argv)
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			set wb to open workbook workbook file name hfs
			set ws to worksheet "S" of wb
			if byKeys then
				activate
				delay 1
				repeat with i from 1 to n
					select range ("A" & i) of ws
					set t to item i of texts
					set er to ""
					if t is not "" then set er to kb's typeKeys(t)
					set end of refusedOut to er
				end repeat
			else
				set display alerts to false
				repeat with i from 1 to n
					set er to ""
					try
						set formula2 local of range ("A" & i) of ws to (item i of texts)
					on error msg number k
						set er to (k as text) & " " & msg
					end try
					set end of refusedOut to er
				end repeat
				set display alerts to true
			end if
			calculate
			repeat with i from 1 to n
				set rg to range ("A" & i) of ws
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
				if withEdit then
					set end of cellsOut to {tx as text, number format of rg, number format local of rg, pc as text, vv, formula of rg, formula local of rg}
				else
					set end of cellsOut to {tx as text, number format of rg, number format local of rg, pc as text, vv}
				end if
			end repeat
			if withEdit then
				-- a bar not read (Excel's window on another desktop) leaves no book open
				try
					set bars to my formulaBars(ws, n, name of wb)
				on error msg number k
					close wb saving no
					error msg number k
				end try
				repeat with i from 1 to n
					set end of item i of cellsOut to item i of bars
				end repeat
			end if
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{refusedOut, cellsOut} forKeys:{"refused", "cells"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run

-- what the formula bar shows for each of the cells A1 to An of the sheet `ws` of the book `bookName`, each selected in
-- turn with Excel in front: the description of the text area of the book's window ("Formula Bar. A,1 <text>")
on formulaBars(ws, n, bookName)
	-- the window lists its formula bar only with Excel in front
	repeat 10 times
		tell application "System Events" to set fp to name of first application process whose frontmost is true
		if fp is "Microsoft Excel" then exit repeat
		tell application "System Events" to set frontmost of process "Microsoft Excel" to true
		tell application "Microsoft Excel" to activate
		delay 0.5
	end repeat
	if fp is not "Microsoft Excel" then error "Excel is not in front: " & fp
	delay 1
	-- the book's window, whose name is the book's (another window may be in front of it)
	set bar to missing value
	set seen to ""
	tell application "System Events" to tell process "Microsoft Excel"
		repeat with w in windows
			set seen to seen & (name of w) & "; "
			if name of w is bookName or name of w starts with (text 1 thru -6 of bookName) then
				set els to entire contents of w
				repeat with e in els
					try
						if role of e is "AXTextArea" then
							set bar to contents of e
							exit repeat
						end if
					end try
				end repeat
			end if
			if bar is not missing value then exit repeat
		end repeat
	end tell
	if bar is missing value then error "no formula bar in the window of " & bookName & " among " & seen
	set out to {}
	repeat with i from 1 to n
		tell application "Microsoft Excel" to select range ("A" & i) of ws
		delay 0.3
		-- the bar follows the selection with a lag: read until two reads agree
		tell application "System Events" to set d to description of bar
		repeat 10 times
			delay 0.3
			tell application "System Events" to set again to description of bar
			if again is d then exit repeat
			set d to again
		end repeat
		set end of out to d
	end repeat
	return out
end formulaBars

-- the handlers of excel-keys.applescript (its path `p`) as a script object
on keysLib(p)
	set src to (current application's NSString's stringWithContentsOfFile:p encoding:4 |error|:(missing value)) as text
	return run script ("script K" & linefeed & src & linefeed & "end script" & linefeed & "return K")
end keysLib
