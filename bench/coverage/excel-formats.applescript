-- Open a book and read the number format of cells of its first sheet as Excel shows it, then close it unsaved.
-- Arguments: the book's path (inside Excel's sandbox), the cells' addresses separated by commas.
-- Output: JSON [format, ...] in the order of the addresses.
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
				set end of out to (number format of range (ad as text) of ws)
			end repeat
			close wb saving no
		end timeout
	end tell
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:out options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
