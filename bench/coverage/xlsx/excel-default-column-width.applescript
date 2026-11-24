-- The width and displayed value of A1 and B1 in the opened book, then the book saved as xlsx.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set sourcePath to (POSIX file (item 1 of argv)) as text
	set savedPath to (POSIX file (item 2 of argv)) as text
	set savedName to item 3 of argv
	set answer to {}
	tell application "Microsoft Excel"
		set wb to open workbook workbook file name sourcePath update links do not update links
		set ws to worksheet "S" of wb
		repeat with cellAddress in {"A1", "B1"}
			set cellRange to range cellAddress of ws
			set end of answer to {width of cellRange, column width of cellRange, string value of cellRange}
		end repeat
		save workbook as wb filename savedPath file format Excel XML file format
		close workbook savedName saving no
	end tell
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:answer options:0 |error|:(reference)
	if d is missing value then error "Excel's width data is not JSON"
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
