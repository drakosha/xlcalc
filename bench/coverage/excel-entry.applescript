-- Open an entry part's book, type each case's text into its cell (column A of the sheet "S", row by row) as Excel 365
-- takes typed text in the user's language (`formula2 local`), alerts off, then read each cell: what it shows (`string
-- value`), its format codes (`number format`, `number format local`), its prefix character and its value (`value`, a
-- number only, not read from a cell that shows an error); save the book as xlsx under another name and close it.
-- With a fifth argument "keys" the texts are typed on the keyboard instead (System Events, Excel in front, checked
-- before each text; an empty text is not typed), each followed by Return; the window Excel shows for a formula it does
-- not take is answered (`answerDialog`) and its text is the text's refusal. With a last argument "edit" each cell's
-- `formula` and `formula local` are read too: the text the formula bar opens the cell with.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the texts as
-- a JSON array of strings, "keys" and the path of excel-keys.applescript, or nothing; then "edit", or nothing.
-- Output: JSON {"refused": [Excel's error for each text, "" when it took it], "cells": [[text, code, local, prefix, value(,
-- formula, formula local)], ...]}, the value "" for a cell whose value is no number (or a date), or that shows an error.
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
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{refusedOut, cellsOut} forKeys:{"refused", "cells"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run

-- the handlers of excel-keys.applescript (its path `p`) as a script object
on keysLib(p)
	set src to (current application's NSString's stringWithContentsOfFile:p encoding:4 |error|:(missing value)) as text
	return run script ("script K" & linefeed & src & linefeed & "end script" & linefeed & "return K")
end keysLib
