-- Opens a measurement book in Excel, recalculates it fully, reads columns B, C, D of the case sheet
-- (type code, number, text) in three calls and closes it without saving. Output: JSON [[B], [C], [D]], and with
-- a third argument `formula2` also column A's formulas as Excel 365 shows them (`@` where it intersects).
-- Arguments: the book's path (inside Excel's sandbox), the number of rows, `formula2` (optional).
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set n to item 2 of argv
	set withF2 to (count of argv) > 2
	set hfs to (POSIX file p) as text
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			set wb to open workbook workbook file name hfs
			calculate full
			set ws to worksheet "C" of wb
			set t to value of range ("B1:B" & n) of ws
			set nums to value of range ("C1:C" & n) of ws
			set strs to value of range ("D1:D" & n) of ws
			set cols to {t, nums, strs}
			if withF2 then set end of cols to formula2 of range ("A1:A" & n) of ws
			close wb saving no
		end timeout
	end tell
	set arr to current application's NSArray's arrayWithArray:cols
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:arr options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
