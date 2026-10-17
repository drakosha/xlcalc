-- Opens a book as it is, computes it fully, adds a sheet whose rows read the given cells (type code, number, text:
-- READ_COLUMNS of book.ts, -5 for a blank cell), reads that sheet and closes the book without saving. Output: JSON
-- [[A], [B], [C]]. Arguments: the book's path (inside Excel's sandbox), the number of cells, the cells' references
-- one per line ('Sheet 1'!B5).
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set n to item 2 of argv
	set refs to paragraphs of (item 3 of argv)
	set hfs to (POSIX file p) as text
	set probes to {}
	repeat with r in refs
		set x to r as text
		set end of probes to {"=IF(ISBLANK(" & x & "),-5,IF(ISERROR(" & x & "),ERROR.TYPE(" & x & "),IF(ISNUMBER(" & x & "),0,IF(ISLOGICAL(" & x & "),IF(" & x & ",-1,-2),IF(ISTEXT(" & x & "),-3,-4)))))", "=IF(ISNUMBER(" & x & ")," & x & ",0)", "=IF(ISTEXT(" & x & ")," & x & ",\"\")"}
	end repeat
	tell application "Microsoft Excel"
		with timeout of 170 seconds
			set wb to open workbook workbook file name hfs
			calculate full
			set ws to make new worksheet at end of wb
			-- a list of rows set on a range at once leaves it empty: cell by cell
			repeat with i from 1 to count of probes
				set {fa, fb, fc} to item i of probes
				set formula of range ("A" & i) of ws to fa
				set formula of range ("B" & i) of ws to fb
				set formula of range ("C" & i) of ws to fc
			end repeat
			calculate
			set t to value of range ("A1:A" & n) of ws
			set nums to value of range ("B1:B" & n) of ws
			set strs to value of range ("C1:C" & n) of ws
			close wb saving no
		end timeout
	end tell
	set arr to current application's NSArray's arrayWithArray:{t, nums, strs}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:arr options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
