-- Open a book, activate a sheet, open Review > Workbook Statistics with Excel in front (Accessibility: the tab and the
-- button pressed by their names), read every text of the window it opens, close the window (OK) and the book unsaved.
-- Arguments: the book's path (inside Excel's sandbox), the sheet's name.
-- Output: the window's texts in its order, one a line: label and value of each row, then the section's title.
use AppleScript version "2.4"
use scripting additions

on run argv
	set p to item 1 of argv
	set sheetName to item 2 of argv
	set hfs to (POSIX file p) as text
	tell application "Microsoft Excel"
		with timeout of 90 seconds
			set wb to open workbook workbook file name hfs
			activate object (worksheet sheetName of wb)
			set bookName to name of wb
		end timeout
	end tell
	tell application "System Events" to set frontmost of process "Microsoft Excel" to true
	tell application "Microsoft Excel" to activate
	delay 1.5
	-- the Review tab, then its button
	-- pressed again while the window has not come (the ribbon laid out again under a press)
	repeat 3 times
		my pressNamed({"AXRadioButton", "AXTab"}, "Review")
		delay 1
		my pressNamed({"AXButton", "AXMenuButton"}, "Workbook Statistics")
		if my opened() then exit repeat
	end repeat
	set out to {}
	tell application "System Events" to tell process "Microsoft Excel"
		set w to window "Workbook Statistics"
		set els to entire contents of w
		repeat with e in els
			try
				if role of e is "AXStaticText" then set end of out to (value of e as text)
			end try
		end repeat
		click button "OK" of w
	end tell
	delay 1
	tell application "Microsoft Excel"
		close wb saving no
	end tell
	set AppleScript's text item delimiters to linefeed
	return out as text
end run

-- whether the window has come within 5 seconds
on opened()
	repeat 10 times
		tell application "System Events" to if exists window "Workbook Statistics" of process "Microsoft Excel" then return true
		delay 0.5
	end repeat
	return false
end opened

-- an element of Excel's windows of one of `roles` pressed by its name, description or title (the list of a window's
-- elements taken once, as excel-ui.applescript's pressButton takes it)
on pressNamed(roles, n)
	repeat 10 times
		tell application "System Events" to set frontmost of process "Microsoft Excel" to true
		delay 0.5
		tell application "System Events" to tell process "Microsoft Excel"
			repeat with w in windows
				set els to entire contents of w
				repeat with e in els
					try
						if (role of e) is in roles and ((name of e as text) is n or (description of e as text) is n or (title of e as text) is n) then
							perform action "AXPress" of e
							return
						end if
					end try
				end repeat
			end repeat
		end tell
	end repeat
	error "no " & n
end pressNamed
