-- Open a table case's book, compute it fully, take the case's steps (Excel recalculating after each), then read: the
-- rectangles' values through mirror formulas written now on the mirror sheets (so they point at the cells where the
-- steps left them; they never give an error), the rectangles' formulas as Excel shows them (and in the user's
-- language with `local`), and the tables of every sheet. Close without saving. Alerts are off during the steps.
-- Arguments: book path (inside Excel's sandbox), steps as lines "kind<TAB>args..." (see below), rectangles as lines
-- "sheet<TAB>range<TAB>mirror range<TAB>type formula<TAB>number formula<TAB>text formula" (the mirror formulas over the
-- rectangle's top-left cell), "1" to read formulas in the user's language too (else empty), the names of the mirror
-- sheets of the types, the numbers and the texts, separated by tabs.
-- Steps: e sheet range n|s|f|l|x value (a number, a text, a formula as Excel 365 enters it, one in the user's language, cleared); insr, delr, insc,
-- delc sheet range (the entire rows or columns of the range); rename sheet table name; deltable sheet table; convert
-- sheet table; copy sheet range sheet range; cut sheet range sheet cell; renamesheet sheet name; delsheet sheet;
-- name scope name formula (a defined name added, global with an empty scope).
-- Output: JSON {"steps": [the error of each step, "" when Excel took it], "rects": [[types, numbers, texts,
-- formulas, local formulas or ""], ...], "tables": [[sheet, name, range, headers shown, totals shown, [column names]], ...],
-- "names": [the visible defined names, a sheet's own as Sheet!name]}.
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
	set mirrors to my splitTab(item 5 of argv)
	set hfs to (POSIX file p) as text
	set stepsOut to {}
	set out to {}
	set tablesOut to {}
	tell application "Microsoft Excel"
		with timeout of 90 seconds
			-- a book Excel refuses to open answers -50 here, with or without its repair window
			try
				set wb to open workbook workbook file name hfs
			on error msg number n
				error "the book did not open: " & msg number n
			end try
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
						else if k is "copy" then
							copy range (range (item 3 of parts) of worksheet (item 2 of parts) of wb) destination (range (item 5 of parts) of worksheet (item 4 of parts) of wb)
						else if k is "cut" then
							cut range (range (item 3 of parts) of worksheet (item 2 of parts) of wb) destination of cut (range (item 5 of parts) of worksheet (item 4 of parts) of wb)
						else if k is "renamesheet" then
							set name of worksheet (item 2 of parts) of wb to (item 3 of parts)
						else if k is "delsheet" then
							delete worksheet (item 2 of parts) of wb
						else if k is "name" then
							if item 2 of parts is "" then
								make new named item at wb with properties {name:(item 3 of parts), references:(item 4 of parts)}
							else
								make new named item at worksheet (item 2 of parts) of wb with properties {name:(item 3 of parts), references:(item 4 of parts)}
							end if
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
					repeat with i from 1 to 3
						set formula of range (item 3 of parts) of worksheet (item i of mirrors) of wb to (item (3 + i) of parts)
					end repeat
				end if
			end repeat
			calculate
			repeat with ln in rects
				if (length of ln) > 0 then
					set parts to my splitTab(ln)
					set mr to item 3 of parts
					set t to value of range mr of worksheet (item 1 of mirrors) of wb
					set nums to value of range mr of worksheet (item 2 of mirrors) of wb
					set strs to value of range mr of worksheet (item 3 of mirrors) of wb
					set rg to range (item 2 of parts) of worksheet (item 1 of parts) of wb
					set f to formula of rg
					set fl to ""
					if withLocal then set fl to formula local of rg
					set end of out to {t, nums, strs, f, fl}
				end if
			end repeat
			repeat with i from 1 to (count of worksheets of wb)
				set ws to worksheet i of wb
				if name of ws is not in mirrors then
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
			-- by index: the elements of `named items of wb` do not resolve
			set namesOut to {}
			repeat with i from 1 to (count of named items of wb)
				set nm to named item i of wb
				if visible of nm then set end of namesOut to (name of nm)
			end repeat
			close wb saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{stepsOut, out, tablesOut, namesOut} forKeys:{"steps", "rects", "tables", "names"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run

-- a number from its text with a decimal point, whatever the system's locale (`as number` reads the locale's separator)
on num(s)
	return ((current application's NSString's stringWithString:s)'s doubleValue()) as real
end num
