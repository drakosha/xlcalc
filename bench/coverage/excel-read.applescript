-- Opens a measurement book in Excel, recalculates it fully, reads columns B, C, D of the case sheet
-- (type code, number, text) in three calls and closes it without saving. Output: JSON [[B], [C], [D]].
-- Arguments: the book's path (inside Excel's sandbox), the number of rows.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set n to item 2 of argv
	set hfs to (POSIX file p) as text
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			set wb to open workbook workbook file name hfs
			calculate full
			set ws to worksheet "C" of wb
			set t to value of range ("B1:B" & n) of ws
			set nums to value of range ("C1:C" & n) of ws
			set strs to value of range ("D1:D" & n) of ws
			close wb saving no
		end timeout
	end tell
	set arr to current application's NSArray's arrayWithArray:{t, nums, strs}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:arr options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
