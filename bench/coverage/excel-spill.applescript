-- Open a spill case's book, compute it fully, apply the case's edits (a value typed into a cell, a formula
-- entered as Excel 365 enters it, or the cells cleared, Excel recalculating after each), read its rectangles and
-- close without saving. An edit Excel refuses (part of an old-kind array) leaves its error and changes nothing.
-- For every rectangle: the values of its mirrors on the mirror sheets (formulas that never give an error), the
-- formulas of the rectangle itself as Excel shows them, and with `withF2` as Excel 365 shows them.
-- Arguments: book path (inside Excel's sandbox), rectangles as lines "sheet<TAB>range<TAB>mirror range",
-- edits as lines "sheet<TAB>range<TAB>kind<TAB>value" (kind n a number, s a text, f a formula, x clear; empty for
-- none), "1" to read formula2 too (else empty), the names of the mirror sheets of the types, the numbers and the
-- texts, separated by tabs.
-- Output: JSON {"rects": [[types, numbers, texts, formulas, formula2 or ""], ...] in the order of the rectangles,
-- "edits": [the error of each edit, "" when Excel took it]}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set rects to paragraphs of item 2 of argv
	set edits to paragraphs of item 3 of argv
	set withF2 to item 4 of argv is "1"
	set AppleScript's text item delimiters to (character id 9)
	set mirrors to text items of item 5 of argv
	set AppleScript's text item delimiters to ""
	set hfs to (POSIX file p) as text
	set out to {}
	set editsOut to {}
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
					set er to ""
					set display alerts to false
					try
						if item 3 of parts is "x" then
							clear contents rg
						else if item 3 of parts is "n" then
							set value of rg to ((item 4 of parts) as number)
						else if item 3 of parts is "f" then
							set formula2 of rg to (item 4 of parts)
						else
							set value of rg to (item 4 of parts)
						end if
					on error msg number k
						set er to (k as text) & " " & msg
					end try
					set display alerts to true
					set end of editsOut to er
					calculate
				end if
			end repeat
			repeat with ln in rects
				set AppleScript's text item delimiters to (character id 9)
				set parts to text items of ln
				set AppleScript's text item delimiters to ""
				set mr to item 3 of parts
				set t to value of range mr of worksheet (item 1 of mirrors) of wb
				set nums to value of range mr of worksheet (item 2 of mirrors) of wb
				set strs to value of range mr of worksheet (item 3 of mirrors) of wb
				set rg to range (item 2 of parts) of worksheet (item 1 of parts) of wb
				set f to formula of rg
				set f2 to ""
				if withF2 then set f2 to formula2 of rg
				set end of out to {t, nums, strs, f, f2}
			end repeat
			close wb saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{out, editsOut} forKeys:{"rects", "edits"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
