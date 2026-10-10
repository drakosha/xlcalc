-- Open a book, make rules of data validation and outlines through Excel's object model, apply one structure operation,
-- read every sheet's rules (each cell of the window that has one, with that cell selected: a rule's formulas are relative
-- to the active cell), whether each such cell's value meets its rule, the outline levels, hidden rows and columns and
-- collapsed summaries of the window and the sheet's summary settings, then save the book as xlsx under another name.
-- Alerts are off around the operation; an operation Excel refuses comes back as `error`, not a dialog.
-- Arguments: the book's path and the path to save to (inside Excel's sandbox), the saved file's name, what to make as JSON
-- ([["rule", sheet, range, first cell, type, alert, operator, formula1, formula2, ignore blank, dropdown, show input,
-- show error, input title, input message, error title, error message], ["group", sheet, "rows"|"columns", range],
-- ["collapse", sheet, "rows"|"columns", cell], ["levels", sheet, row levels, column levels], ["summary", sheet, below,
-- right]]: flags "1", "0" or "" for Excel's own), the operation as JSON ([op, sheet, range, target sheet or new name,
-- target], or []), the window (`A1:J30`).
-- Output: JSON {"error": text, "sheets": [[name, [[address, type, operator, formula1, formula2, ignore blank, dropdown,
-- show input, show error, alert style, input title, input message, error title, error message, value]], [[row, level,
-- hidden, collapsed]], [[column, level, hidden, collapsed]], [summary row, summary column]]]}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on parsed(t)
	set d to (current application's NSString's stringWithString:t)'s dataUsingEncoding:4
	return (current application's NSJSONSerialization's JSONObjectWithData:d options:0 |error|:(missing value)) as list
end parsed

-- "" for a missing value
on txt(v)
	if v is missing value then return ""
	return v as text
end txt

on dvType(t)
	tell application "Microsoft Excel"
		if t is "whole" then return validate whole number
		if t is "decimal" then return validate decimal
		if t is "list" then return validate list
		if t is "date" then return validated date
		if t is "time" then return validate time
		if t is "textLength" then return validate text length
		if t is "custom" then return validate custom
		return validate input only
	end tell
end dvType

on dvAlert(t)
	tell application "Microsoft Excel"
		if t is "warning" then return valid alert warning
		if t is "information" then return valid alert information
		return valid alert stop
	end tell
end dvAlert

on dvOperator(t)
	tell application "Microsoft Excel"
		if t is "notBetween" then return operator not between
		if t is "equal" then return operator equal
		if t is "notEqual" then return operator not equal
		if t is "greaterThan" then return operator greater
		if t is "lessThan" then return operator less
		if t is "greaterThanOrEqual" then return operator greater equal
		if t is "lessThanOrEqual" then return operator less equal
		return operator between
	end tell
end dvOperator

-- a column's letters by its number
on letters(n)
	set out to ""
	repeat while n > 0
		set k to (n - 1) mod 26
		set out to (character id (65 + k)) & out
		set n to (n - 1 - k) div 26
	end repeat
	return out
end letters

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set makes to parsed(item 4 of argv)
	set opArgs to parsed(item 5 of argv)
	set win to item 6 of argv
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set opError to ""
	set readError to ""
	set sheetsOut to {}
	tell application "Microsoft Excel"
		with timeout of 120 seconds
			set wb to open workbook workbook file name hfs
			set stage to "make"
			try
				repeat with m in makes
					set mk to item 1 of m
					set ws to worksheet (item 2 of m) of wb
					activate object ws
					if mk is "rule" then
						-- relative references of a rule's formulas are relative to the active cell
						select range (item 4 of m) of ws
						set v to validation of range (item 3 of m) of ws
						set tt to my dvType(item 5 of m)
						if (item 5 of m) is "any" then
							add data validation v type tt
						else if (item 9 of m) is "" then
							add data validation v type tt alert style (my dvAlert(item 6 of m)) operator (my dvOperator(item 7 of m)) formula1 (item 8 of m)
						else
							add data validation v type tt alert style (my dvAlert(item 6 of m)) operator (my dvOperator(item 7 of m)) formula1 (item 8 of m) formula2 (item 9 of m)
						end if
						if (item 10 of m) is not "" then set ignore blank of v to ((item 10 of m) is "1")
						if (item 11 of m) is not "" then set in cell dropdown of v to ((item 11 of m) is "1")
						if (item 12 of m) is not "" then set show input of v to ((item 12 of m) is "1")
						if (item 13 of m) is not "" then set show error of v to ((item 13 of m) is "1")
						if (item 14 of m) is not "" then set input title of v to (item 14 of m)
						if (item 15 of m) is not "" then set input message of v to (item 15 of m)
						if (item 16 of m) is not "" then set error title of v to (item 16 of m)
						if (item 17 of m) is not "" then set error message of v to (item 17 of m)
					else if mk is "group" then
						if (item 3 of m) is "rows" then
							group (entire row of range (item 4 of m) of ws)
						else
							group (entire column of range (item 4 of m) of ws)
						end if
					else if mk is "collapse" then
						if (item 3 of m) is "rows" then
							set show detail of (entire row of range (item 4 of m) of ws) to false
						else
							set show detail of (entire column of range (item 4 of m) of ws) to false
						end if
					else if mk is "levels" then
						show levels (outline object of ws) row levels (item 3 of m) column levels (item 4 of m)
					else if mk is "summary" then
						if (item 3 of m) is "1" then set summary row of (outline object of ws) to summary below
						if (item 3 of m) is "0" then set summary row of (outline object of ws) to summary above
						if (item 4 of m) is "1" then set summary column of (outline object of ws) to summary on right
						if (item 4 of m) is "0" then set summary column of (outline object of ws) to summary on left
					end if
				end repeat
				if (count of opArgs) > 0 then
					set stage to "op"
					set op to item 1 of opArgs
					set argSheet to item 2 of opArgs
					set argRange to item 3 of opArgs
					set argTarget to item 4 of opArgs
					set argTo to item 5 of opArgs
					set display alerts to false
					if argSheet is not "" then set ws to worksheet argSheet of wb
					if op is "insertRows" then
						insert into range (entire row of (range argRange of ws)) shift shift down
					else if op is "deleteRows" then
						delete range (entire row of (range argRange of ws)) shift shift up
					else if op is "insertColumns" then
						insert into range (entire column of (range argRange of ws)) shift shift to right
					else if op is "deleteColumns" then
						delete range (entire column of (range argRange of ws)) shift shift to left
					else if op is "moveRange" then
						cut range (range argRange of ws) destination of cut (range argTo of worksheet argTarget of wb)
					else if op is "copyRange" then
						copy range (range argRange of ws) destination (range argTo of worksheet argTarget of wb)
					else if op is "renameSheet" then
						set name of ws to argTarget
					else if op is "deleteSheet" then
						delete ws
					else
						error "unknown op " & op
					end if
					set display alerts to true
				end if
			on error msg number k
				set display alerts to true
				set opError to stage & ": " & (k as text) & " " & msg
			end try
			set stage to "sheets"
			try
				repeat with i from 1 to (count of worksheets of wb)
					set w to worksheet i of wb
					-- special cells and selection work on the active sheet
					activate object w
					set stage to "rules"
					set rulesOut to {}
					set addrText to ""
					-- no rule in the window: a parameter error
					try
						set addrText to get address (special cells (range win of w) type cell type all validation)
					end try
					if addrText is not "" then
						set AppleScript's text item delimiters to ","
						set areaList to text items of addrText
						set AppleScript's text item delimiters to ""
						repeat with ad in areaList
							set ar to range (ad as text) of w
							repeat with j from 1 to (count of cells of ar)
								set ca to get address (cell j of ar)
								select range ca of w
								set v to validation of (range ca of w)
								set f1 to ""
								set f2 to ""
								try
									set f1 to my txt(formula1 of v)
								end try
								try
									set f2 to my txt(formula2 of v)
								end try
								set opText to ""
								try
									set opText to (validation operator of v) as text
								end try
								set end of rulesOut to {ca, (validation type of v) as text, opText, f1, f2, ignore blank of v, in cell dropdown of v, show input of v, show error of v, (alert style of v) as text, my txt(input title of v), my txt(input message of v), my txt(error title of v), my txt(error message of v), value of v}
							end repeat
						end repeat
					end if
					set stage to "outline"
					set winRange to range win of w
					set firstRow to first row index of winRange
					set lastRow to firstRow + (count of rows of winRange) - 1
					set firstCol to first column index of winRange
					set lastCol to firstCol + (count of columns of winRange) - 1
					set rowsOut to {}
					repeat with r from firstRow to lastRow
						set er to entire row of range ("A" & r) of w
						set lv to outline level of er
						set hd to hidden of er
						set cl to false
						if summary of er then set cl to not (show detail of er)
						if lv > 1 or hd or cl then set end of rowsOut to {r, lv - 1, hd, cl}
					end repeat
					set colsOut to {}
					repeat with c from firstCol to lastCol
						set ec to entire column of range ((my letters(c)) & "1") of w
						set lv to outline level of ec
						set hd to hidden of ec
						set cl to false
						if summary of ec then set cl to not (show detail of ec)
						if lv > 1 or hd or cl then set end of colsOut to {my letters(c), lv - 1, hd, cl}
					end repeat
					set ol to outline object of w
					set end of sheetsOut to {name of w, rulesOut, rowsOut, colsOut, {(summary row of ol) as text, (summary column of ol) as text}}
				end repeat
			on error msg number k
				set readError to stage & ": " & (k as text) & " " & msg
			end try
			if readError is "" then
				activate object (worksheet 1 of wb)
				select range "A1" of (worksheet 1 of wb)
				save workbook as wb filename out file format Excel XML file format
				close workbook savedName saving no
			else
				close wb saving no
			end if
		end timeout
	end tell
	if readError is not "" then error "reading: " & readError
	set d0 to current application's NSDictionary's dictionaryWithObjects:{opError, sheetsOut} forKeys:{"error", "sheets"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
