-- Open a ui case's book (sheet "S"), compute it fully and take the case's action with Excel in front, on the keyboard
-- (System Events, the handlers of excel-keys.applescript): "bars" selects each cell in turn and reads the formula bar,
-- then presses F2 and reads the bar and the element in focus, and Escape; "paste" puts the text on the clipboard as plain
-- text and presses Cmd+V (the V key: any layout) with the cell selected; "fill" selects the range, types the text and presses Control+Return;
-- "edit" selects the cell, presses F2, then Left `left` times and Backspace `back` times (40 without "keys": all it
-- opened with), types the text (`type`, or `text` without "keys") and presses Return, reading the bar as "bars" does
-- and once more before Return; "pick" selects the cell, opens its dropdown list (Option+Down, the first item marked),
-- presses Down `item` - 1 times and Return;
-- "ribbon" selects each cell in turn and reads the ribbon's Number Format box, "choose" applies an item of the box's list
-- to each cell (Accessibility: the box's button pressed, then the item) and reads the box after; both give A1 the text
-- format as a control cell, selected before each cell, and a set leaves A1 and A2 empty.
-- "point" selects the cell and takes each step in turn, a text typed (the first begins the cell's editing over what it
-- holds, Enter mode) or a key pressed by its code with Shift or Command held; after each the formula bar's description
-- is read and a window Excel shows answered; Escape ends what editing is left without writing it.
-- Windows Excel shows for a paste or a fill are answered (`answerDialog`). Then read each cell: what it shows (`string
-- value`), its format codes, prefix character, value (a number only, not read from a cell that shows an error), formula
-- (`formula`, `formula local`, `formula2`, `formula2 local`) and whether it is part of a formula of the old kind (`has
-- array`, with its `formula array`); each column's `column width`, `width` and `use standard width`; save the book as
-- xlsx under another name and close it.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the action as
-- JSON {"do": "bars" | "paste" | "fill" | "edit" | "ribbon" | "choose" | "point" | "none", "cells": [...], "items": [...], "at": cell,
-- "range": range, "text": text, "keys": {"left", "back", "type"}, "steps": [{"type"} | {"key", "shift", "cmd"}]}, the cells to
-- read and the columns as JSON arrays, the path of excel-keys.applescript. "none" takes no action: the book is only read.
-- Output: JSON {"dialog": text, "bars": [[[description, value] selected, [description, value] in editing, focus, dialog,
-- [description, value] typed (an edit's)], ...], "ribbon": [[item, box], ...], "point": [[bar's description, dialog], ...], "cells": [[text, code, local, prefix, value, formula, formula local, formula2, formula2
-- local, has array, formula array], ...], "columns": [[column width, points, standard], ...]}.
use AppleScript version "2.4"
use framework "Foundation"
use framework "AppKit"
use scripting additions

-- the ribbon's Number Format box, once found (`findBox`)
global numberBox

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set act to my fromJSON(item 4 of argv)
	set reads to my fromJSON(item 5 of argv) as list
	set cols to my fromJSON(item 6 of argv) as list
	set kb to my keysLib(item 7 of argv)
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set todo to (act's objectForKey:"do") as text
	set said to ""
	set barsOut to {}
	set ribbonOut to {}
	set pointOut to {}
	set cellsOut to {}
	set colsOut to {}
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			set wb to open workbook workbook file name hfs
			set ws to worksheet "S" of wb
			set bookName to name of wb
			calculate full
		end timeout
	end tell
	try
		if todo is "bars" then
			set barsOut to my bars(ws, (act's objectForKey:"cells") as list, bookName, kb)
		else if todo is "paste" then
			my toClipboard((act's objectForKey:"text") as text)
			tell application "Microsoft Excel" to select range ((act's objectForKey:"at") as text) of ws
			kb's toFront()
			-- the V key, whatever the layout types on it, as a person pastes
			tell application "System Events" to key code 9 using command down
			delay 1.5
			set said to kb's answerDialog()
		else if todo is "fill" then
			tell application "Microsoft Excel" to select range ((act's objectForKey:"range") as text) of ws
			set said to kb's typeFill((act's objectForKey:"text") as text)
		else if todo is "edit" then
			set barsOut to my edit(ws, act, bookName, kb)
		else if todo is "pick" then
			set said to my pick(ws, (act's objectForKey:"at") as text, (act's objectForKey:"item") as integer, kb)
		else if todo is "ribbon" then
			set ribbonOut to my ribbon(ws, (act's objectForKey:"cells") as list, {}, bookName, kb)
		else if todo is "choose" then
			set ribbonOut to my ribbon(ws, (act's objectForKey:"cells") as list, (act's objectForKey:"items") as list, bookName, kb)
		else if todo is "point" then
			set pointOut to my pointSteps(ws, act, bookName, kb)
		end if
	on error msg number k
		-- a step that failed (Excel's window on another desktop) leaves no book open: an editing it left is ended first,
		-- and the step's own error is the one reported
		try
			tell application "System Events" to key code 53
			delay 0.3
			tell application "Microsoft Excel" to close wb saving no
		end try
		error msg number k
	end try
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			calculate
			repeat with a in reads
				set rg to range (a as text) of ws
				set tx to string value of rg
				if tx is missing value then set tx to ""
				set pc to prefix character of rg
				if pc is missing value then set pc to ""
				-- an error's value is not read (a range's errors crashed Excel); a number too wide for the column shows ## and is read
				set vv to ""
				if tx does not start with "#" or tx starts with "##" then
					try
						set vv to value of rg
					end try
					-- a number alone: a date comes as a date object, which JSON has no place for
					if class of vv is not real and class of vv is not integer then set vv to ""
				end if
				set ha to has array of rg
				set fa to ""
				if ha then set fa to formula array of rg
				set end of cellsOut to {tx as text, number format of rg, number format local of rg, pc as text, vv, formula of rg, formula local of rg, formula2 of rg, formula2 local of rg, ha, fa}
			end repeat
			repeat with c in cols
				set firstCell to range ((c as text) & "1") of ws
				set end of colsOut to {column width of firstCell, width of firstCell, use standard width of firstCell}
			end repeat
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{said, barsOut, ribbonOut, pointOut, cellsOut, colsOut} forKeys:{"dialog", "bars", "ribbon", "point", "cells", "columns"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run

-- each cell of `cells` selected in turn with Excel in front: the formula bar then, and once F2 began the cell's editing,
-- with the element in focus; Escape ends the editing, and a window Excel shows meanwhile is answered
on bars(ws, cells, bookName, kb)
	set bar to kb's formulaBar(bookName)
	set out to {}
	repeat with a in cells
		tell application "Microsoft Excel" to select range (a as text) of ws
		set onSelect to kb's barText(bar)
		kb's toFront()
		tell application "System Events" to key code 120
		delay 0.5
		set onEdit to kb's barText(bar)
		set inFocus to kb's focusedText()
		tell application "System Events" to key code 53
		delay 0.5
		set end of out to {onSelect, onEdit, inFocus, kb's answerDialog()}
	end repeat
	return out
end bars

-- the cell selected and its editing begun with F2, the cursor at the end of what it opens with: Left and Backspace
-- pressed, the text typed, Return; the formula bar as `bars` reads it, and once the keys were typed
on edit(ws, act, bookName, kb)
	set k to act's objectForKey:"keys"
	set {l, b, t} to {0, 40, (act's objectForKey:"text") as text}
	if k is not missing value then set {l, b, t} to {(k's objectForKey:"left") as integer, (k's objectForKey:"back") as integer, (k's objectForKey:"type") as text}
	set bar to kb's formulaBar(bookName)
	tell application "Microsoft Excel" to select range ((act's objectForKey:"at") as text) of ws
	set onSelect to kb's barText(bar)
	kb's toFront()
	tell application "System Events" to key code 120
	delay 0.5
	set onEdit to kb's barText(bar)
	set inFocus to kb's focusedText()
	kb's toFront()
	tell application "System Events"
		repeat l times
			key code 123
		end repeat
		repeat b times
			key code 51
		end repeat
		keystroke t
	end tell
	set onTyped to kb's barText(bar)
	kb's toFront()
	tell application "System Events" to keystroke return
	delay 1
	return {{onSelect, onEdit, inFocus, kb's answerDialog(), onTyped}}
end edit

-- the cell selected; each step a text typed or a key pressed by its code (any layout) with its modifiers, then the
-- formula bar's description read and a window Excel showed answered; Escape ends the editing left
on pointSteps(ws, act, bookName, kb)
	set bar to kb's formulaBar(bookName)
	tell application "Microsoft Excel" to select range ((act's objectForKey:"at") as text) of ws
	set out to {}
	-- the steps kept an NSArray of NSDictionary: `as list` would make them records, which take no objectForKey
	set allSteps to act's objectForKey:"steps"
	repeat with i from 0 to ((allSteps's |count|()) - 1)
		set stp to (allSteps's objectAtIndex:i)
		kb's toFront()
		set t to (stp's objectForKey:"type")
		if t is not missing value then
			tell application "System Events" to keystroke (t as text)
		else
			set c to my keyCode((stp's objectForKey:"key") as text)
			set withShift to (stp's objectForKey:"shift") is not missing value
			set withCmd to (stp's objectForKey:"cmd") is not missing value
			tell application "System Events"
				if withShift and withCmd then
					key code c using {shift down, command down}
				else if withShift then
					key code c using shift down
				else if withCmd then
					key code c using command down
				else
					key code c
				end if
			end tell
		end if
		delay 0.5
		set saidNow to my answerAlerts()
		set end of out to {item 1 of kb's barText(bar), saidNow}
	end repeat
	kb's toFront()
	tell application "System Events" to key code 53
	delay 0.3
	return out
end pointSteps

-- The alerts Excel shows during a case of `point`, answered as a person answers them (No, OK, Cancel) and their texts;
-- a window without buttons (the tip of a function's arguments) is left, and no Escape follows: the editing goes on.
on answerAlerts()
	set said to ""
	repeat 5 times
		set found to false
		tell application "System Events" to tell process "Microsoft Excel"
			-- a window that closes meanwhile (a tip) is skipped: the list is taken once and each window read in a try
			set ws to {}
			try
				set ws to every window
			end try
			repeat with w in ws
				set sr to ""
				set names to {}
				try
					set sr to subrole of w
					set names to name of every button of w
				end try
				if sr is not "AXStandardWindow" and names is not {} then
					set tx to ""
					repeat with x in (static texts of w)
						try
							set v to value of x
							if v is not missing value then set tx to tx & (v as text) & " "
						end try
					end repeat
					set said to said & tx & "| "
					repeat with b in {"No", "Нет", "OK", "ОК", "Cancel", "Отмена"}
						if names contains (b as text) then
							click button (b as text) of w
							set found to true
							exit repeat
						end if
					end repeat
					if not found then error "an Excel alert without No, OK or Cancel: " & tx
					exit repeat
				end if
			end repeat
		end tell
		if not found then exit repeat
		delay 1
	end repeat
	if said is "" then return ""
	return "dialog: " & said
end answerAlerts

-- a key of a case of `point` by its code on the Mac's keyboard
on keyCode(k)
	if k is "left" then return 123
	if k is "right" then return 124
	if k is "down" then return 125
	if k is "up" then return 126
	if k is "home" then return 115
	if k is "pagedown" then return 121
	if k is "f2" then return 120
	if k is "return" then return 36
	if k is "escape" then return 53
	error "no key " & k
end keyCode

-- the cell selected, its dropdown list opened (Option+Down, key codes: any layout; an empty cell's list opens with its
-- first item marked), the item reached by Down and taken by Return; the windows Excel shows then are answered
on pick(ws, a, n, kb)
	tell application "Microsoft Excel" to select range a of ws
	kb's toFront()
	tell application "System Events" to key code 125 using option down
	delay 0.8
	tell application "System Events"
		repeat (n - 1) times
			key code 125
			delay 0.15
		end repeat
		key code 36
	end tell
	delay 1
	return kb's answerDialog()
end pick

-- each cell selected in turn, an item of the Number Format list applied to it when `picks` gives one (its place in the
-- list): the item's name (empty for none) and what the box shows then. Before each cell the control cell A1, given the
-- text format, is selected and the box must show the control's category (`control`): a box that does not follow the
-- selection is looked for again, then stops the case.
on ribbon(ws, cells, picks, bookName, kb)
	tell application "Microsoft Excel" to set number format of range "A1" of ws to "@"
	my findBox(bookName)
	set ctl to my control(ws, "", bookName, kb)
	set out to {}
	repeat with i from 1 to count of cells
		set a to (item i of cells) as text
		set ctl to my control(ws, ctl, bookName, kb)
		tell application "Microsoft Excel" to select range a of ws
		set shown to my boxAfter(ctl, bookName)
		set nm to ""
		if picks is not {} then
			set k to (item i of picks) as integer
			set m to my openList(bookName, a)
			tell application "System Events"
				set nm to name of menu item k of m
				perform action "AXPress" of menu item k of m
			end tell
			set shown to my boxAfter(shown, bookName)
		end if
		set end of out to {nm, shown}
	end repeat
	return out
end ribbon

-- The control cell A1 selected, until the box shows `ctl` (the text category): its category. Empty `ctl`, at the start:
-- the empty A2 selected first, A1's category is what the box changes to from A2's. A box that keeps another value for 3
-- seconds is looked for again, three times at most.
on control(ws, ctl, bookName, kb)
	if ctl is "" then
		kb's toFront()
		tell application "Microsoft Excel" to select range "A2" of ws
		delay 1
		set was to my boxNow(bookName)
		tell application "Microsoft Excel" to select range "A1" of ws
		return my boxAfter(was, bookName)
	end if
	repeat 3 times
		kb's toFront()
		tell application "Microsoft Excel" to select range "A1" of ws
		repeat 30 times
			delay 0.1
			if my boxNow(bookName) is ctl then return ctl
		end repeat
		my findBox(bookName)
	end repeat
	error "the Number Format box does not follow the selection: it shows " & my boxNow(bookName) & ", not " & ctl
end control

-- the box's list opened by its button: its menu, once it is there
on openList(bookName, a)
	my boxNow(bookName)
	tell application "System Events"
		perform action "AXPress" of button 1 of numberBox
		repeat 20 times
			delay 0.2
			try
				set m to menu 1 of numberBox
				delay 0.3
				return m
			end try
		end repeat
	end tell
	error "the Number Format list did not open at " & a
end openList

-- The ribbon's Number Format box of the window of the book `bookName`, by its description in English or Russian: the
-- window lists it only with Excel in front, so it is looked for again until it does; an element of the window's
-- contents that has gone stale is passed over. Kept in `numberBox`.
on findBox(bookName)
	global numberBox
	repeat 10 times
		tell application "System Events" to set frontmost of process "Microsoft Excel" to true
		tell application "Microsoft Excel" to activate
		delay 1.5
		set numberBox to missing value
		try
			tell application "System Events" to tell process "Microsoft Excel"
				repeat with w in windows
					set wn to name of w
					if wn is bookName or bookName starts with wn then
						set els to entire contents of w
						repeat with e in els
							try
								if role of e is "AXComboBox" and description of e is in {"Number Format", "Формат номера"} then
									set numberBox to contents of e
									exit repeat
								end if
							end try
						end repeat
					end if
				end repeat
			end tell
		end try
		if numberBox is not missing value then return numberBox
	end repeat
	error "no Number Format box in the window of " & bookName
end findBox

-- What the box shows now; a box that has gone stale (the ribbon laid out again) is looked for again.
on boxNow(bookName)
	global numberBox
	try
		tell application "System Events" to set v to value of numberBox
		if v is not missing value then return v as text
	end try
	my findBox(bookName)
	tell application "System Events" to return (value of numberBox) as text
end boxNow

-- What the box shows once it has followed a selection or an item applied: the box changes from what it showed before
-- (`was`) within 2 seconds (0.3 s seen; the control shows it follows), or the cell's category is the same; then two reads agree.
on boxAfter(was, bookName)
	set v to was
	repeat 20 times
		delay 0.1
		set v to my boxNow(bookName)
		if v is not was then exit repeat
	end repeat
	repeat 10 times
		delay 0.3
		set again to my boxNow(bookName)
		if again is v then exit repeat
		set v to again
	end repeat
	return v
end boxAfter

-- the text on the general pasteboard as plain text, nothing else with it
on toClipboard(t)
	set pb to current application's NSPasteboard's generalPasteboard()
	pb's clearContents()
	pb's setString:t forType:(current application's NSPasteboardTypeString)
end toClipboard

on fromJSON(s)
	set js to current application's NSString's stringWithString:s
	return current application's NSJSONSerialization's JSONObjectWithData:(js's dataUsingEncoding:4) options:0 |error|:(missing value)
end fromJSON

-- the handlers of excel-keys.applescript (its path `p`) as a script object
on keysLib(p)
	set src to (current application's NSString's stringWithContentsOfFile:p encoding:4 |error|:(missing value)) as text
	return run script ("script K" & linefeed & src & linefeed & "end script" & linefeed & "return K")
end keysLib
