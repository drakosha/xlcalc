-- Open a book, copy a range to the clipboard, paste it with Paste Special at a target, then save the book as xlsx under
-- another name. Alerts are off around the paste; a paste Excel refuses comes back as `error`, not a dialog. All Merging
-- Conditional Formats, which AppleScript does not name, goes through the Paste Special window (Command+Control+V, Excel in
-- front): its kind, operation, Skip Blanks and Transpose clicked through System Events and OK; an alert after it is read
-- as `error` and dismissed. The Format Painter (painter) takes Excel to the front: the source selected, the Home tab's
-- Format button clicked (System Events), the selection given (the target and one more line) and Shift with the key
-- shrinking it onto the target, as the painter applies at a change of the selection by the keyboard, not by the object
-- model; with more targets the button is double-clicked (a mouse event), which keeps it on, each target taken the same
-- way and Escape turning it off. A source in another book: `source.xlsx` beside the book, opened, copied from and closed.
-- Arguments: the book's path and the path to save to (inside Excel's sandbox), the saved file's name, the paste as JSON
-- ([source sheet, source range, target sheet, target range, what (all, values, formulas, formats, validation,
-- formulasAndNumberFormats, valuesAndNumberFormats, columnWidths, allExceptBorders, notes, allUsingSourceTheme,
-- allMergingConditionalFormats; copy: a plain copy to the target, Range.Copy with a destination; painter: the Format
-- Painter), transpose "1" or "0", the operation (none, add, subtract, multiply, divide), skip blanks "1" or "0", the
-- painter's key (left or up), cells of the target sheet whose formula local is read after the paste, separated by commas,
-- the painter's further targets ([sheet, selection, key]), "1" when the source is in source.xlsx]).
-- Output: JSON {"error": text, "local": [formula local of each cell]}; an error of the script's own steps (the window did
-- not open, the painter did not turn on or off) begins with "harness: ", which is no answer of Excel's.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on parsed(t)
	set d to (current application's NSString's stringWithString:t)'s dataUsingEncoding:4
	return (current application's NSJSONSerialization's JSONObjectWithData:d options:0 |error|:(missing value)) as list
end parsed

on pasteType(t)
	tell application "Microsoft Excel"
		if t is "values" then return paste values
		if t is "formulas" then return paste formulas
		if t is "formats" then return paste formats
		if t is "validation" then return paste validation
		if t is "formulasAndNumberFormats" then return paste formulas and number formats
		if t is "valuesAndNumberFormats" then return paste values and number formats
		if t is "columnWidths" then return paste column widths
		if t is "allExceptBorders" then return paste all except borders
		if t is "notes" then return paste comments
		if t is "allUsingSourceTheme" then return paste all using source theme
		if t is "all" then return paste all
	end tell
	error "no paste type " & t
end pasteType

on operationOf(t)
	tell application "Microsoft Excel"
		if t is "add" then return paste special operation add
		if t is "subtract" then return paste special operation subtract
		if t is "multiply" then return paste special operation multiply
		if t is "divide" then return paste special operation divide
		if t is "none" then return paste special operation none
	end tell
	error "no operation " & t
end operationOf

-- The Format button of the Home tab (the Format Painter) in the window of the book `bookName`, Excel in front.
on painterButton(bookName)
	repeat 5 times
		tell application "System Events" to set frontmost of process "Microsoft Excel" to true
		tell application "Microsoft Excel" to activate
		delay 1.5
		tell application "System Events" to tell process "Microsoft Excel"
			repeat with w in windows
				if (name of w as text) is bookName or bookName starts with (name of w as text) then
					-- a list: a repeat over the reference itself fails
					set els to entire contents of w
					repeat with e in els
						try
							if role of e is "AXCheckBox" and name of e is "Format" then return e
						end try
					end repeat
				end if
			end repeat
		end tell
	end repeat
	error "harness: no Format Painter in the window of " & bookName
end painterButton

-- Excel in front, a key with Shift: the painter applies at the change of the selection it makes.
on shiftKey(k)
	tell application "System Events"
		if k is "up" then
			key code 126 using shift down
		else
			key code 123 using shift down
		end if
	end tell
end shiftKey

-- A double click of the left button at the middle of an element, as a mouse event (System Events clicks once).
on doubleClick(e)
	tell application "System Events"
		set {x, y} to position of e
		set {w, h} to size of e
	end tell
	set js to "ObjC.import('CoreGraphics'); var p = $.CGPointMake(" & (x + w / 2) & "," & (y + h / 2) & "); [[1,1],[2,1],[1,2],[2,2]].forEach(function (t) { var e = $.CGEventCreateMouseEvent(null, t[0], p, 0); $.CGEventSetIntegerValueField(e, 1, t[1]); $.CGEventPost(0, e); delay(0.05) })"
	do shell script "osascript -l JavaScript -e " & quoted form of js
end doubleClick

-- The source selected on its sheet, the Format Painter on (double-clicked for more targets), each target selected with
-- one more line and the key shrinking it; Escape turns a double-clicked painter off.
on paint(wb, ws, wt, op, bookName)
	set more to item 11 of op
	tell application "Microsoft Excel"
		activate object ws
		select range (item 2 of op) of ws
	end tell
	set b to my painterButton(bookName)
	if (count of more) is 0 then
		tell application "System Events" to click b
	else
		my doubleClick(b)
	end if
	delay 0.8
	tell application "System Events" to if (value of b as integer) is 0 then error "harness: the Format Painter did not turn on"
	tell application "Microsoft Excel"
		activate object wt
		select range (item 4 of op) of wt
	end tell
	delay 0.5
	my shiftKey(item 9 of op)
	delay 1
	repeat with t in more
		tell application "System Events" to if (value of b as integer) is 0 then error "harness: the Format Painter turned off before " & (item 2 of t)
		tell application "Microsoft Excel"
			activate object (worksheet (item 1 of t) of wb)
			select range (item 2 of t) of (worksheet (item 1 of t) of wb)
		end tell
		delay 0.5
		my shiftKey(item 3 of t)
		delay 1
	end repeat
	if (count of more) is not 0 then
		tell application "System Events" to key code 53
		delay 0.8
	end if
	tell application "System Events" to if (value of b as integer) is not 0 then error "harness: the Format Painter stayed on"
end paint

-- The window's radio button or check box (`kind`, its role) named `n`, clicked when its value is not `v`.
on setControl(w, kind, n, v)
	tell application "System Events" to tell process "Microsoft Excel"
		-- a list: a repeat over the reference itself fails
		set els to entire contents of w
		repeat with e in els
			try
				if (role of e as text) is kind and (name of e as text) is n then
					if (value of e as integer) is not v then click e
					return
				end if
			end try
		end repeat
	end tell
	error "harness: no " & n & " in the Paste Special window"
end setControl

-- The window of Excel named `n`, waited for up to `secs` seconds; missing value if it did not come.
on excelWindow(n, secs)
	repeat (secs * 5) times
		tell application "System Events" to tell process "Microsoft Excel"
			if exists window n then return window n
		end tell
		delay 0.2
	end repeat
	return missing value
end excelWindow

-- The text of Excel's windows that are not books, joined; the window answered with Return (its default button).
on alertText()
	set t to ""
	tell application "System Events" to tell process "Microsoft Excel"
		repeat with w in windows
			if (subrole of w as text) is not "AXStandardWindow" then
				repeat with e in (entire contents of w)
					try
						if role of e is "AXStaticText" then set t to t & (value of e as text) & " "
					end try
				end repeat
				perform action "AXRaise" of w
				key code 36
			end if
		end repeat
	end tell
	return t
end alertText

-- Paste Special through its window: the target selected, the window opened, the kind, operation, Skip Blanks and
-- Transpose set and OK; an alert after it comes back as the error.
on pasteInWindow(wt, op)
	tell application "Microsoft Excel"
		activate object wt
		select range (item 4 of op) of wt
		activate
	end tell
	-- the keys again where they were lost, Excel in front each time
	set w to missing value
	repeat 3 times
		tell application "System Events" to set frontmost of process "Microsoft Excel" to true
		delay 1
		tell application "System Events" to keystroke "v" using {command down, control down}
		set w to my excelWindow("Paste Special", 5)
		if w is not missing value then exit repeat
	end repeat
	if w is missing value then error "harness: no Paste Special window"
	set ops to {{"none", "None"}, {"add", "Add"}, {"subtract", "Subtract"}, {"multiply", "Multiply"}, {"divide", "Divide"}}
	repeat with o in ops
		if (item 1 of o) is (item 7 of op) then my setControl(w, "AXRadioButton", item 2 of o, 1)
	end repeat
	my setControl(w, "AXRadioButton", "All, merge conditional formats", 1)
	my setControl(w, "AXCheckBox", "Skip Blanks", (item 8 of op) as integer)
	my setControl(w, "AXCheckBox", "Transpose", (item 6 of op) as integer)
	tell application "System Events" to tell process "Microsoft Excel" to click button "OK" of w
	delay 1.5
	set t to my alertText()
	if t is not "" then error t
end pasteInWindow

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set op to parsed(item 4 of argv)
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set opError to ""
	tell application "Microsoft Excel"
		with timeout of 120 seconds
			set wb to open workbook workbook file name hfs
			set wf to wb
			if (item 12 of op) is "1" then
				set src to (POSIX file ((do shell script "dirname " & quoted form of p) & "/source.xlsx")) as text
				set wf to open workbook workbook file name src
			end if
			set stage to "copy"
			try
				set display alerts to false
				set ws to worksheet (item 1 of op) of wf
				activate object ws
				set wt to worksheet (item 3 of op) of wb
				if (item 5 of op) is "painter" then
					set stage to "paste"
					my paint(wb, ws, wt, op, name of wb)
				else if (item 5 of op) is "copy" then
					set stage to "paste"
					copy range (range (item 2 of op) of ws) destination (range (item 4 of op) of wt)
				else if (item 5 of op) is "allMergingConditionalFormats" then
					copy range (range (item 2 of op) of ws)
					set stage to "paste"
					my pasteInWindow(wt, op)
				else
					copy range (range (item 2 of op) of ws)
					set stage to "paste"
					activate object wt
					paste special (range (item 4 of op) of wt) what (my pasteType(item 5 of op)) operation (my operationOf(item 7 of op)) skip blanks ((item 8 of op) is "1") transpose ((item 6 of op) is "1")
				end if
				set display alerts to true
			on error msg number k
				set display alerts to true
				set opError to stage & ": " & (k as text) & " " & msg
			end try
			set cut copy mode to false
			set locals to {}
			if (item 10 of op) is not "" then
				set AppleScript's text item delimiters to ","
				set cellNames to text items of (item 10 of op)
				set AppleScript's text item delimiters to ""
				repeat with a in cellNames
					set end of locals to (formula local of range (a as text) of worksheet (item 3 of op) of wb) as text
				end repeat
			end if
			activate object (worksheet 1 of wb)
			select range "A1" of (worksheet 1 of wb)
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
			if (item 12 of op) is "1" then close workbook "source.xlsx" saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{opError, locals} forKeys:{"error", "local"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
