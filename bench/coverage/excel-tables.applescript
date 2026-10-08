-- Open a table case's book, compute it fully, take the case's steps (Excel recalculating after each), then read: the
-- rectangles' values through mirror formulas written now on the sheets WT, WN, WX (so they point at the cells where
-- the steps left them; they never give an error), the rectangles' formulas as Excel shows them (and in the user's
-- language with `local`), and the tables of every sheet. Close without saving. Alerts are off during the steps.
-- Arguments: book path (inside Excel's sandbox), steps as lines "kind<TAB>args..." (see below), rectangles as lines
-- "sheet<TAB>range<TAB>mirror range", "1" to read formulas in the user's language too (else empty).
-- Steps: e sheet range n|s|f|l|x value (a number, a text, a formula as Excel 365 enters it, one in the user's language, cleared); insr, delr, insc,
-- delc sheet range (the entire rows or columns of the range); rename sheet table name; deltable sheet table; convert
-- sheet table; totals sheet table 1|0; copy sheet range sheet range; cut sheet range sheet cell; renamesheet sheet
-- name; delsheet sheet.
-- Output: JSON {"steps": [the error of each step, "" when Excel took it], "rects": [[types, numbers, texts,
-- formulas, local formulas or ""], ...], "tables": [[sheet, name, range, headers shown, totals shown, [column names]], ...]}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on splitTab(ln)
	set AppleScript's text item delimiters to (character id 9)
	set parts to text items of ln
	set AppleScript's text item delimiters to ""
	return parts
end splitTab

on run argv
	set p to item 1 of argv
	set steps to paragraphs of item 2 of argv
	set rects to paragraphs of item 3 of argv
	set withLocal to item 4 of argv is "1"
	set hfs to (POSIX file p) as text
	set stepsOut to {}
	set out to {}
	set tablesOut to {}
	set mirror to {"IF(ISBLANK(#),-5,IF(ISERROR(#),ERROR.TYPE(#),IF(ISNUMBER(#),0,IF(ISLOGICAL(#),IF(#,-1,-2),IF(ISTEXT(#),-3,-4)))))", "IF(ISNUMBER(#),#,0)", "IF(ISTEXT(#),#,\"\")"}
	tell application "Microsoft Excel"
		with timeout of 90 seconds
			set wb to open workbook workbook file name hfs
			calculate full
			repeat with ln in steps
				if (length of ln) > 0 then
					set parts to my splitTab(ln)
					set k to item 1 of parts
					set er to ""
					set display alerts to false
					try
						if k is "e" then
							set rg to range (item 3 of parts) of worksheet (item 2 of parts) of wb
							if item 4 of parts is "x" then
								clear contents rg
							else if item 4 of parts is "n" then
								set value of rg to (my num(item 5 of parts))
							else if item 4 of parts is "f" then
								set formula2 of rg to (item 5 of parts)
							else if item 4 of parts is "l" then
								set formula local of rg to (item 5 of parts)
							else
								set value of rg to (item 5 of parts)
							end if
						else if k is "insr" then
							insert into range (entire row of (range (item 3 of parts) of worksheet (item 2 of parts) of wb)) shift shift down
						else if k is "delr" then
							delete range (entire row of (range (item 3 of parts) of worksheet (item 2 of parts) of wb)) shift shift up
						else if k is "insc" then
							insert into range (entire column of (range (item 3 of parts) of worksheet (item 2 of parts) of wb)) shift shift to right
						else if k is "delc" then
							delete range (entire column of (range (item 3 of parts) of worksheet (item 2 of parts) of wb)) shift shift to left
						else if k is "rename" then
							set name of list object (item 3 of parts) of worksheet (item 2 of parts) of wb to (item 4 of parts)
						else if k is "deltable" then
							delete (list object (item 3 of parts) of worksheet (item 2 of parts) of wb)
						else if k is "convert" then
							convert to range (list object (item 3 of parts) of worksheet (item 2 of parts) of wb)
						else if k is "totals" then
							set total of list object (item 3 of parts) of worksheet (item 2 of parts) of wb to ((item 4 of parts) is "1")
						else if k is "copy" then
							copy range (range (item 3 of parts) of worksheet (item 2 of parts) of wb) destination (range (item 5 of parts) of worksheet (item 4 of parts) of wb)
						else if k is "cut" then
							cut range (range (item 3 of parts) of worksheet (item 2 of parts) of wb) destination of cut (range (item 5 of parts) of worksheet (item 4 of parts) of wb)
						else if k is "renamesheet" then
							set name of worksheet (item 2 of parts) of wb to (item 3 of parts)
						else if k is "delsheet" then
							delete worksheet (item 2 of parts) of wb
						else
							error "unknown step " & k
						end if
					on error msg number n
						set er to (n as text) & " " & msg
					end try
					set display alerts to true
					set end of stepsOut to er
					calculate
				end if
			end repeat
			-- the mirrors, written now, relative to each rectangle's top-left cell
			repeat with ln in rects
				if (length of ln) > 0 then
					set parts to my splitTab(ln)
					set AppleScript's text item delimiters to ":"
					set tl to text item 1 of (item 2 of parts)
					set AppleScript's text item delimiters to ""
					set src to "'" & my quoted(item 1 of parts) & "'!" & tl
					repeat with i from 1 to 3
						set AppleScript's text item delimiters to "#"
						set bits to text items of (item i of mirror)
						set AppleScript's text item delimiters to src
						set f to "=" & (bits as text)
						set AppleScript's text item delimiters to ""
						set formula of range (item 3 of parts) of worksheet (item i of {"WT", "WN", "WX"}) of wb to f
					end repeat
				end if
			end repeat
			calculate
			repeat with ln in rects
				if (length of ln) > 0 then
					set parts to my splitTab(ln)
					set mr to item 3 of parts
					set t to value of range mr of worksheet "WT" of wb
					set nums to value of range mr of worksheet "WN" of wb
					set strs to value of range mr of worksheet "WX" of wb
					set rg to range (item 2 of parts) of worksheet (item 1 of parts) of wb
					set f to formula of rg
					set fl to ""
					if withLocal then set fl to formula local of rg
					set end of out to {t, nums, strs, f, fl}
				end if
			end repeat
			repeat with i from 1 to (count of worksheets of wb)
				set ws to worksheet i of wb
				if name of ws is not in {"WT", "WN", "WX"} then
					repeat with j from 1 to (count of list objects of ws)
						set lo to list object j of ws
						set cols to {}
						repeat with c from 1 to (count of list columns of lo)
							set end of cols to name of list column c of lo
						end repeat
						set end of tablesOut to {name of ws, name of lo, get address of (range object of lo), show headers of lo, total of lo, cols}
					end repeat
				end if
			end repeat
			close wb saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{stepsOut, out, tablesOut} forKeys:{"steps", "rects", "tables"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run

-- a number from its text with a decimal point, whatever the system's locale (`as number` reads the locale's separator)
on num(s)
	return ((current application's NSString's stringWithString:s)'s doubleValue()) as real
end num

-- a sheet name with its apostrophes doubled, for a quoted reference
on quoted(s)
	set AppleScript's text item delimiters to "'"
	set bits to text items of s
	set AppleScript's text item delimiters to "''"
	set r to bits as text
	set AppleScript's text item delimiters to ""
	return r
end quoted
