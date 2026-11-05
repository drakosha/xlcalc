-- Open a book and read the fill of cells of its first sheet and their font's colour as Excel shows them, then close it unsaved.
-- Arguments: the book's path (inside Excel's sandbox), the cells' addresses separated by commas.
-- Output: JSON [[[r, g, b], pattern, [r, g, b]], ...] in the order of the addresses, the pattern as Excel names it.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set AppleScript's text item delimiters to ","
	set addrs to text items of (item 2 of argv)
	set AppleScript's text item delimiters to ""
	set hfs to (POSIX file p) as text
	set out to {}
	tell application "Microsoft Excel"
		with timeout of 120 seconds
			set wb to open workbook workbook file name hfs
			set ws to worksheet 1 of wb
			repeat with ad in addrs
				set c to range (ad as text) of ws
				set end of out to {color of interior object of c, (pattern of interior object of c) as text, color of font object of c}
			end repeat
			close wb saving no
		end timeout
	end tell
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:out options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
