-- Open a spill case's book, compute it fully, apply the case's edits (a value typed into a cell or the cell
-- cleared, Excel recalculating after each), read its rectangles and close without saving. For every
-- rectangle: the values of its mirrors on the sheets WT, WN and WX (formulas that never give an error) and
-- the formulas of the rectangle itself as Excel shows them.
-- Arguments: book path (inside Excel's sandbox), rectangles as lines "sheet<TAB>range<TAB>mirror range",
-- edits as lines "sheet<TAB>cell<TAB>kind<TAB>value" (kind n a number, s a text, x clear; empty for none).
-- Output: JSON [[types, numbers, texts, formulas], ...] in the order of the rectangles.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set rects to paragraphs of item 2 of argv
	set edits to paragraphs of item 3 of argv
	set hfs to (POSIX file p) as text
	set out to {}
	tell application "Microsoft Excel"
		with timeout of 90 seconds
			set wb to open workbook workbook file name hfs
			calculate full
			repeat with ln in edits
				if (length of ln) > 0 then
					set AppleScript's text item delimiters to (character id 9)
					set parts to text items of ln
					set AppleScript's text item delimiters to ""
					set rg to range (item 2 of parts) of worksheet (item 1 of parts) of wb
					if item 3 of parts is "x" then
						clear contents rg
					else if item 3 of parts is "n" then
						set value of rg to ((item 4 of parts) as number)
					else
						set value of rg to (item 4 of parts)
					end if
					calculate
				end if
			end repeat
			repeat with ln in rects
				set AppleScript's text item delimiters to (character id 9)
				set parts to text items of ln
				set AppleScript's text item delimiters to ""
				set mr to item 3 of parts
				set t to value of range mr of worksheet "WT" of wb
				set nums to value of range mr of worksheet "WN" of wb
				set strs to value of range mr of worksheet "WX" of wb
				set f to formula of range (item 2 of parts) of worksheet (item 1 of parts) of wb
				set end of out to {t, nums, strs, f}
			end repeat
			close wb saving no
		end timeout
	end tell
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:out options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
