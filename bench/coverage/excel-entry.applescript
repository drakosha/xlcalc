-- Open an entry part's book, type each case's text into its cell (column A of the sheet "S", row by row) as Excel 365
-- takes typed text in the user's language (`formula2 local`), alerts off, then read each cell: what it shows (`string
-- value`), its format codes (`number format`, `number format local`) and its prefix character; save the book as xlsx
-- under another name and close it.
-- With a fifth argument "keys" the texts are typed on the keyboard instead (System Events, Excel in front, checked
-- before each text; an empty text is not typed), each followed by Return.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the texts as
-- a JSON array of strings, "keys" or nothing.
-- Output: JSON {"refused": [Excel's error for each text, "" when it took it], "cells": [[text, code, local, prefix], ...]}.
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
					if t is not "" then my typeKeys(t)
					set end of refusedOut to ""
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
				set end of cellsOut to {tx as text, number format of rg, number format local of rg, pc as text}
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

-- a text typed into the selected cell, then Return; Excel must be the application in front
on typeKeys(t)
	tell application "System Events"
		set fp to name of first application process whose frontmost is true
		if fp is not "Microsoft Excel" then error "Excel is not in front: " & fp
		keystroke t
		delay 0.2
		keystroke return
	end tell
	delay 0.3
end typeKeys
