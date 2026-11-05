-- Open a book of the filter-acts sets (sheet "S"), read the sheet, and take each step of the case, reading the sheet
-- after each; a step's `via` says how: through the object model (`autofilter range`, the sort object of the AutoFilter,
-- `autofilter mode`), or as a person does with Excel in front: the header's button opened with Option+Down, its Sort
-- Descending or its values' check boxes clicked, Data > AutoFilter from the menu with the cell or the range selected.
-- Then save the book as xlsx under another name and close it.
-- Arguments: the book's path, the path to save to (both inside Excel's sandbox), the saved file's name, the case as JSON
-- (`actSpec`), the path of excel-keys.applescript. Output: JSON {"opened": state, "after": [state, ...], "dialog": text}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

property spec : missing value
property kb : missing value
property said : ""
property listed : missing value

on run argv
	set hfs to (POSIX file (item 1 of argv)) as text
	set out to (POSIX file (item 2 of argv)) as text
	set savedName to item 3 of argv
	set spec to my fromJSON(item 4 of argv)
	set kb to my keysLib(item 5 of argv)
	set steps to spec's objectForKey:"steps"
	set afterOut to {}
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			set wb to open workbook workbook file name hfs
			set ws to worksheet "S" of wb
			calculate full
		end timeout
	end tell
	try
		set opened to my stateNow(ws)
		repeat with i from 0 to ((steps's |count|()) as integer) - 1
			my takeStep(ws, steps's objectAtIndex:i)
			set end of afterOut to my stateNow(ws)
		end repeat
	on error msg number k
		try
			tell application "System Events" to key code 53
			delay 0.3
			tell application "Microsoft Excel" to close wb saving no
		end try
		error msg number k
	end try
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{opened, afterOut, said} forKeys:{"opened", "after", "dialog"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run

on takeStep(ws, stp)
	set listed to missing value
	set keysOf to (stp's allKeys()) as list
	if keysOf contains "criteria" then
		set v to (stp's objectForKey:"criteria") as text
		set col to (stp's objectForKey:"column") as text
		if ((stp's objectForKey:"via") as text) is "button" then
			my criteriaButton(ws, col, v)
		else
			set rg to my filterRange(ws)
			tell application "Microsoft Excel" to autofilter range rg field (my fieldOf(rg, col)) criteria1 v
		end if
	else if keysOf contains "sort" then
		set col to (stp's objectForKey:"sort") as text
		if ((stp's objectForKey:"via") as text) is "button" then
			my sortButton(ws, col)
		else
			my sortModel(ws, col)
		end if
	else if keysOf contains "off" then
		if ((stp's objectForKey:"off") as text) is "button" then
			tell application "Microsoft Excel" to select range "A1" of ws
			my pressFilter()
		else
			tell application "Microsoft Excel" to set autofilter mode of ws to false
		end if
	else if keysOf contains "over" then
		set a to (stp's objectForKey:"over") as text
		set how to (stp's objectForKey:"via") as text
		if how is "button" then
			tell application "Microsoft Excel" to select range a of ws
			my pressFilter()
		else if how is "field" then
			tell application "Microsoft Excel" to autofilter range (range a of ws) field 1
		else
			tell application "Microsoft Excel" to autofilter range (range a of ws)
		end if
	else if keysOf contains "set" then
		tell application "Microsoft Excel" to set formula of range ((stp's objectForKey:"at") as text) of ws to ((stp's objectForKey:"set") as text)
	else if keysOf contains "totals" then
		tell application "Microsoft Excel"
			set total of list object 1 of ws to true
			set totals calculation of list column 3 of list object 1 of ws to sum totals calc
		end tell
	end if
end takeStep

-- the AutoFilter object of the sheet, or of its table
on filterOf(ws)
	tell application "Microsoft Excel"
		if ((spec's objectForKey:"table") as boolean) then return autofilter object of list object 1 of ws
		return autofilter object of ws
	end tell
end filterOf

on filterRange(ws)
	tell application "Microsoft Excel" to return range object of (my filterOf(ws))
end filterRange

-- a column's number from the first of the range
on fieldOf(rg, col)
	tell application "Microsoft Excel" to return (first column index of (range (col & "1") of worksheet object of rg)) - (first column index of rg) + 1
end fieldOf

-- the sort object of the AutoFilter, its key the column over the range it acts on, largest first
on sortModel(ws, col)
	set rg to my filterRange(ws)
	tell application "Microsoft Excel"
		set t to first row index of rg
		set b to t + (count of rows of rg) - 1
		set k to range (col & t & ":" & col & b) of ws
		set so to sort object of (my filterOf(ws))
		set fs to sortfieldset of so
		clear sortfieldset fs
		add sortfield fs key k sorton sort on values order sort descending
		apply sort so
	end tell
end sortModel

-- a column's button opened from its header cell: its floating window
on openButton(ws, col)
	set rg to my filterRange(ws)
	tell application "Microsoft Excel"
		set t to first row index of rg
		select range (col & t) of ws
	end tell
	kb's toFront()
	tell application "System Events" to key code 125 using {option down}
	repeat 20 times
		delay 0.25
		tell application "System Events" to tell process "Microsoft Excel"
			repeat with w in windows
				if subrole of w is "AXFloatingWindow" then return contents of w
			end repeat
		end tell
	end repeat
	error "the button of column " & col & " opened no window"
end openButton

-- the values a button lists, without (Select All)
on listedOf(w)
	set out to {}
	tell application "System Events"
		set els to entire contents of w
		repeat with e in els
			try
				if role of e is "AXCheckBox" and description of e is "checkbox" then
					set t to my trimmed(title of e as text)
					if t is not "Auto Apply" and t is not "(Select All)" then set end of out to t
				end if
			end try
		end repeat
	end tell
	return out
end listedOf

on trimmed(t)
	repeat while t ends with " "
		set t to text 1 thru -2 of t
	end repeat
	return t
end trimmed

-- the first element of a window with this role whose description or title is `name`
on elementOf(w, r, what)
	tell application "System Events"
		set els to entire contents of w
		repeat with e in els
			try
				if role of e is r then
					if (description of e as text) is what then return contents of e
					if my trimmed(title of e as text) is what then return contents of e
				end if
			end try
		end repeat
	end tell
	error "no " & r & " " & what & " in the button's window"
end elementOf

on closeButton(w)
	tell application "System Events" to click my elementOf(w, "AXButton", "close button")
	delay 0.8
end closeButton

on sortButton(ws, col)
	set w to my openButton(ws, col)
	set listed to my listedOf(w)
	tell application "System Events" to click my elementOf(w, "AXRadioButton", "Sort Descending")
	delay 1
	set w to my openAgain(ws, col)
	my closeButton(w)
end sortButton

-- the button's window, if a click left it open; opened again otherwise, to be closed by the same hand
on openAgain(ws, col)
	tell application "System Events" to tell process "Microsoft Excel"
		repeat with w in windows
			if subrole of w is "AXFloatingWindow" then return contents of w
		end repeat
	end tell
	return my openButton(ws, col)
end openAgain

on criteriaButton(ws, col, v)
	set w to my openButton(ws, col)
	set listed to my listedOf(w)
	tell application "System Events"
		click my elementOf(w, "AXCheckBox", "(Select All)")
		delay 0.5
		click my elementOf(w, "AXCheckBox", v)
		delay 0.5
		set b to my elementOf(w, "AXButton", "Apply Filter")
		if enabled of b then click b
	end tell
	delay 1
	my closeButton(my openAgain(ws, col))
end criteriaButton

-- Data > AutoFilter from the menu bar, the button of the Data tab; a window Excel shows for it answered
on pressFilter()
	kb's toFront()
	tell application "System Events" to tell process "Microsoft Excel" to click menu item "AutoFilter" of menu "Data" of menu bar 1
	delay 1.2
	set said to said & kb's answerDialog()
end pressFilter

-- the sheet as read, as a list the step turns into its fields
on stateNow(ws)
	set rowsRead to (spec's objectForKey:"rows") as integer
	set watched to (spec's objectForKey:"watch") as list
	set totalLabel to (spec's objectForKey:"label") as text
	set hid to {}
	set heights to {}
	set col to {}
	set tot to missing value
	tell application "Microsoft Excel"
		set rangeNow to ""
		if ((spec's objectForKey:"table") as boolean) then
			set rangeNow to get address of (range object of (autofilter object of list object 1 of ws))
		else if autofilter mode of ws then
			set rangeNow to get address of (range object of (autofilter object of ws))
		end if
		set mode to filter mode of ws
		set std to standard height of ws
		repeat with r from 1 to rowsRead
			set rg to range ("A" & r) of ws
			if hidden of (entire row of rg) then set end of hid to r
			if watched contains r then set end of heights to {r, row height of rg}
			if r > 1 then set end of col to string value of range ("C" & r) of ws
			if tot is missing value and (string value of rg) is totalLabel then set tot to {r, string value of range ("C" & r) of ws, formula of range ("C" & r) of ws}
		end repeat
	end tell
	set AppleScript's text item delimiters to "$"
	set parts to text items of rangeNow
	set AppleScript's text item delimiters to ""
	set rangeNow to parts as text
	if tot is missing value then set tot to {}
	set l to listed
	if l is missing value then set l to {}
	return {rangeNow, mode, hid, heights, std, tot, col, l}
end stateNow

on fromJSON(s)
	set js to current application's NSString's stringWithString:s
	return current application's NSJSONSerialization's JSONObjectWithData:(js's dataUsingEncoding:4) options:0 |error|:(missing value)
end fromJSON

on keysLib(p)
	set src to (current application's NSString's stringWithContentsOfFile:p encoding:4 |error|:(missing value)) as text
	return run script ("script K" & linefeed & src & linefeed & "end script" & linefeed & "return K")
end keysLib
