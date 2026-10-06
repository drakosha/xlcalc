-- Open a display case book, compute it fully, read what each listed cell shows and close without saving.
-- For every cell: its shown text (`string value`, Range.Text) and the colour of its display format's font, which
-- is the colour of the format's section (black for the automatic colour; its colour index reads as missing).
-- Arguments: book path (inside Excel's sandbox), the number of cells n; sheet "L" lists their addresses on
-- sheet "S" in A1:An.
-- Output: JSON [[text, colour], ...] in the order of the addresses, the colour as {r, g, b} ("" when unreadable).
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set n to (item 2 of argv) as integer
	set hfs to (POSIX file p) as text
	set out to {}
	tell application "Microsoft Excel"
		with timeout of 170 seconds
			set wb to open workbook workbook file name hfs
			calculate full
			set ws to worksheet "S" of wb
			if n is 1 then
				set addrs to {value of range "A1" of worksheet "L" of wb}
			else
				set addrs to value of range ("A1:A" & n) of worksheet "L" of wb
			end if
			repeat with ad in addrs
				if class of ad is list then set ad to item 1 of ad
				set rg to range (ad as text) of ws
				set tx to string value of rg
				if tx is missing value then set tx to ""
				set cl to ""
				try
					set cl to color of font object of display format of rg
				end try
				set end of out to {tx as text, cl}
			end repeat
			close wb saving no
		end timeout
	end tell
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:out options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
