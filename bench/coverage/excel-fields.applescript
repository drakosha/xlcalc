-- Open a book, read each worksheet's window with the sheet active (frozen, split, the lines in points, each pane's
-- first row and column, the active pane), save the book as xlsx under another name with its first sheet active, close it.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name.
-- Output: JSON [[name, frozen, split, split horizontal, split vertical, [[scroll row, scroll column], ...], active pane]].
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set sheetsOut to {}
	tell application "Microsoft Excel"
		with timeout of 90 seconds
			set wb to open workbook workbook file name hfs
			repeat with i from 1 to (count of worksheets of wb)
				set ws to worksheet i of wb
				activate object ws
				set aw to active window
				set panesOut to {}
				repeat with j from 1 to (count of panes of aw)
					set pn to pane j of aw
					set end of panesOut to {scroll row of pn, scroll column of pn}
				end repeat
				set end of sheetsOut to {name of ws, freeze panes of aw, split of aw, split horizontal of aw, split vertical of aw, panesOut, entry_index of (active pane of aw)}
			end repeat
			activate object (worksheet 1 of wb)
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:sheetsOut options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
