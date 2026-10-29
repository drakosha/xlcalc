-- Open a filter case's book (sheet "S", an AutoFilter on it), apply the case's criteria to its AutoFilter, then take
-- each step with Excel in front, on the keyboard (System Events, the handlers of excel-keys.applescript): "type" selects
-- the cell, types the text and presses Return; "set" writes the cell's formula through the object model; "paste" puts
-- the text on the clipboard and presses Cmd+V with the cell selected; "copy" selects the range, presses Cmd+C, selects
-- the target, presses Cmd+V and Escape; "fill" selects the range and presses Cmd+D (Fill Down); "clear" selects the
-- range and presses Delete (forward delete, which clears the contents); "undo" presses Cmd+Z; "criteria" filters the
-- range Excel reads (`autofilter range` over `range object of autofilter object`). Keys go by their codes,
-- whatever the layout. After each step the AutoFilter's range is read (`range object of autofilter object`, empty
-- without one) and a window Excel shows is answered. Then save the book as xlsx under another name and close it.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the case as
-- JSON {"criteria": {"field", "value"} (absent without), "filter": range, "steps": [...]}, the path of excel-keys.applescript.
-- Output: JSON {"after": [range, ...], "dialog": text}.
use AppleScript version "2.4"
use framework "Foundation"
use framework "AppKit"
use scripting additions

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set c to my fromJSON(item 4 of argv)
	set kb to my keysLib(item 5 of argv)
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set steps to c's objectForKey:"steps"
	set crit to c's objectForKey:"criteria"
	set said to ""
	set afterOut to {}
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			set wb to open workbook workbook file name hfs
			set ws to worksheet "S" of wb
			calculate full
		end timeout
	end tell
	try
		if crit is not missing value then
			set f to (crit's objectForKey:"field") as integer
			set v to (crit's objectForKey:"value") as text
			tell application "Microsoft Excel" to autofilter range (range ((c's objectForKey:"filter") as text) of ws) field f criteria1 v
		end if
		repeat with i from 0 to ((steps's |count|()) as integer) - 1
			set stp to steps's objectAtIndex:i
			set keysOf to (stp's allKeys()) as list
			if keysOf contains "type" then
				tell application "Microsoft Excel" to select range ((stp's objectForKey:"at") as text) of ws
				set t to (stp's objectForKey:"type") as text
				set said to said & kb's typeKeys(t)
			else if keysOf contains "set" then
				tell application "Microsoft Excel" to set formula of range ((stp's objectForKey:"at") as text) of ws to ((stp's objectForKey:"set") as text)
			else if keysOf contains "paste" then
				my toClipboard((stp's objectForKey:"paste") as text)
				tell application "Microsoft Excel" to select range ((stp's objectForKey:"at") as text) of ws
				set said to said & my keys({9}, {command down}, kb)
			else if keysOf contains "copy" then
				tell application "Microsoft Excel" to select range ((stp's objectForKey:"copy") as text) of ws
				set said to said & my keys({8}, {command down}, kb)
				tell application "Microsoft Excel" to select range ((stp's objectForKey:"to") as text) of ws
				set said to said & my keys({9}, {command down}, kb)
				set said to said & my keys({53}, {}, kb)
			else if keysOf contains "fill" then
				tell application "Microsoft Excel" to select range ((stp's objectForKey:"fill") as text) of ws
				set said to said & my keys({2}, {command down}, kb)
			else if keysOf contains "clear" then
				tell application "Microsoft Excel" to select range ((stp's objectForKey:"clear") as text) of ws
				set said to said & my keys({117}, {}, kb)
			else if keysOf contains "criteria" then
				set k to stp's objectForKey:"criteria"
				tell application "Microsoft Excel" to autofilter range (range object of (autofilter object of ws)) field ((k's objectForKey:"field") as integer) criteria1 ((k's objectForKey:"value") as text)
			else if keysOf contains "undo" then
				set said to said & my keys({6}, {command down}, kb)
			end if
			set end of afterOut to my filterNow(ws)
		end repeat
	on error msg number k
		-- a step that failed leaves no book open: an editing it left is ended first, and the step's own error is the one reported
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
	set d0 to current application's NSDictionary's dictionaryWithObjects:{afterOut, said} forKeys:{"after", "dialog"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run

-- keys pressed by their codes with the modifiers held, Excel in front; the windows Excel shows for them answered: their text
on keys(codes, mods, kb)
	kb's toFront()
	tell application "System Events"
		repeat with k in codes
			key code k using mods
		end repeat
	end tell
	delay 1.2
	return kb's answerDialog()
end keys

-- the AutoFilter's range without its dollar signs, empty without one
on filterNow(ws)
	tell application "Microsoft Excel"
		if not (autofilter mode of ws) then return ""
		set a to get address of (range object of (autofilter object of ws))
	end tell
	set AppleScript's text item delimiters to "$"
	set parts to text items of a
	set AppleScript's text item delimiters to ""
	return parts as text
end filterNow

on toClipboard(t)
	set pb to current application's NSPasteboard's generalPasteboard()
	pb's clearContents()
	pb's setString:t forType:(current application's NSPasteboardTypeString)
end toClipboard

on fromJSON(s)
	set js to current application's NSString's stringWithString:s
	return current application's NSJSONSerialization's JSONObjectWithData:(js's dataUsingEncoding:4) options:0 |error|:(missing value)
end fromJSON

on keysLib(p)
	set src to (current application's NSString's stringWithContentsOfFile:p encoding:4 |error|:(missing value)) as text
	return run script ("script K" & linefeed & src & linefeed & "end script" & linefeed & "return K")
end keysLib
