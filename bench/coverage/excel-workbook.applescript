-- Open a workbook case's book and run its steps: compute (fully or as a recalculation), type a value into a cell
-- or clear it, hide or show rows and columns, filter, read the rectangles; close without saving. Before the steps
-- the application's iteration settings are read (Excel takes them from the book it opens). A read takes, for every
-- rectangle, the values of its mirrors on the sheets WT, WN and WX (formulas that never give an error) and the
-- rectangle's formulas as Excel shows them, and the text each cell to show shows (`string value`, Range.Text).
-- Arguments: book path (inside Excel's sandbox), rectangles as lines "sheet<TAB>range<TAB>mirror range", steps as
-- lines: "full", "rebuild", "calc", "read", "set<TAB>sheet<TAB>cell<TAB>kind<TAB>value" (kind n a number, s a text, x clear),
-- "formula<TAB>sheet<TAB>cell<TAB>=formula" (as one of the old kind), "formula2<TAB>sheet<TAB>cell<TAB>=formula" (as Excel 365 enters it),
-- "formulaarray<TAB>sheet<TAB>range<TAB>=formula" (an array formula of the old kind over the range),
-- "hide<TAB>sheet<TAB>range<TAB>1|0" (the entire rows of the range), "hidecols<TAB>sheet<TAB>range<TAB>1|0",
-- "filter<TAB>sheet<TAB>range<TAB>field<TAB>criterion"; the cells to show as lines "sheet<TAB>address"; the names of
-- the mirror sheets of the types, the numbers and the texts, separated by tabs.
-- Output: JSON {"settings": [iteration, max iterations, max change], "reads": [[[[types, numbers, texts, formulas], ...], [shown text, ...]], ...]}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set rects to paragraphs of item 2 of argv
	set steps to paragraphs of item 3 of argv
	set shows to paragraphs of item 4 of argv
	set AppleScript's text item delimiters to (character id 9)
	set mirrors to text items of item 5 of argv
	set AppleScript's text item delimiters to ""
	set hfs to (POSIX file p) as text
	set reads to {}
	tell application "Microsoft Excel"
		with timeout of 90 seconds
			set wb to open workbook workbook file name hfs
			set settings to {iteration, max iterations, max change}
			repeat with ln in steps
				set AppleScript's text item delimiters to (character id 9)
				set parts to text items of ln
				set AppleScript's text item delimiters to ""
				set k to item 1 of parts
				if k is "full" then
					calculate full
				else if k is "rebuild" then
					calculate full rebuild
				else if k is "calc" then
					calculate
				else if k is "set" then
					set rg to range (item 3 of parts) of worksheet (item 2 of parts) of wb
					if item 4 of parts is "x" then
						clear contents rg
					else if item 4 of parts is "n" then
						set value of rg to ((item 5 of parts) as number)
					else
						set value of rg to (item 5 of parts)
					end if
				else if k is "formula" then
					set formula of range (item 3 of parts) of worksheet (item 2 of parts) of wb to (item 4 of parts)
				else if k is "formula2" then
					set formula2 of range (item 3 of parts) of worksheet (item 2 of parts) of wb to (item 4 of parts)
				else if k is "formulaarray" then
					set formula array of range (item 3 of parts) of worksheet (item 2 of parts) of wb to (item 4 of parts)
				else if k is "hide" then
					set hidden of entire row of (range (item 3 of parts) of worksheet (item 2 of parts) of wb) to ((item 4 of parts) is "1")
				else if k is "hidecols" then
					set hidden of entire column of (range (item 3 of parts) of worksheet (item 2 of parts) of wb) to ((item 4 of parts) is "1")
				else if k is "filter" then
					autofilter range (range (item 3 of parts) of worksheet (item 2 of parts) of wb) field ((item 4 of parts) as integer) criteria1 (item 5 of parts)
				else if k is "read" then
					set snap to {}
					repeat with rl in rects
						set AppleScript's text item delimiters to (character id 9)
						set rp to text items of rl
						set AppleScript's text item delimiters to ""
						set mr to item 3 of rp
						set t to value of range mr of worksheet (item 1 of mirrors) of wb
						set nums to value of range mr of worksheet (item 2 of mirrors) of wb
						set strs to value of range mr of worksheet (item 3 of mirrors) of wb
						set f to formula of range (item 2 of rp) of worksheet (item 1 of rp) of wb
						set end of snap to {t, nums, strs, f}
					end repeat
					set shown to {}
					repeat with sl in shows
						if (length of sl) > 0 then
							set AppleScript's text item delimiters to (character id 9)
							set sp to text items of sl
							set AppleScript's text item delimiters to ""
							set tx to string value of range (item 2 of sp) of worksheet (item 1 of sp) of wb
							if tx is missing value then set tx to ""
							set end of shown to (tx as text)
						end if
					end repeat
					set end of reads to {snap, shown}
				else if (length of k) > 0 then
					error "unknown step " & k
				end if
			end repeat
			close wb saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{settings, reads} forKeys:{"settings", "reads"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
