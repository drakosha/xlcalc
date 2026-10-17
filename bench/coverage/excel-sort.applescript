-- Opens a book, computes it fully, sorts each given range of sheet S by its first column without a header (the Sort
-- command), reads the range's formulas and closes the book without saving. Output: JSON [[formulas of the first
-- range], ...]. Arguments: the book's path (inside Excel's sandbox), the number of ranges, the ranges one per line as
-- "A1:A8 a" (a ascending, d descending).
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set specs to paragraphs of (item 3 of argv)
	set hfs to (POSIX file p) as text
	set out to {}
	-- parsed outside the tell: Excel's own classes would take words such as `word`
	set todo to {}
	repeat with s in specs
		set end of todo to {word 1 of (s as text) & ":" & word 2 of (s as text), word 1 of (s as text), last character of (s as text) is "d"}
	end repeat
	tell application "Microsoft Excel"
		with timeout of 170 seconds
			set wb to open workbook workbook file name hfs
			-- a failed step closes the book before it errs, so no book stays open
			set stage to "calculate"
			try
			calculate full
			set stage to "sheet"
			set ws to worksheet "S" of wb
			-- by index, from a list not named ranges: inside the tell, `contents` and `ranges` are Excel's own
			repeat with i from 1 to count of todo
				set {a, k, desc} to item i of todo
				set stage to "range " & a
				set rg to range a of ws
				set stage to "sort " & a
				if desc then
					sort rg key1 (range k of ws) order1 sort descending header header no
				else
					sort rg key1 (range k of ws) order1 sort ascending header header no
				end if
				set stage to "read " & a
				set end of out to formula of rg
			end repeat
			on error msg number en
				close wb saving no
				error stage & ": " & msg number en
			end try
			close wb saving no
		end timeout
	end tell
	set arr to current application's NSArray's arrayWithArray:out
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:arr options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
