-- Typing on the keyboard, for the steps that type texts as a user does (excel-entry.applescript, excel-fit.applescript,
-- excel-ui.applescript), which load these handlers with `keysLib`: no `use` statements here, as the loader wraps the file
-- in a script object.
-- a text typed into the selected cell, then Return; Excel must be the application in front (it is brought forward for up
-- to 3 seconds). A window Excel shows for it (a formula it does not take) is answered, and Escape leaves the cell as it
-- was: its text is returned.
on typeKeys(t)
	return my typeThen(t, false)
end typeKeys

-- a text typed into the selected cells, then Control+Return, which enters it into each cell of the selection; the windows
-- answered as by typeKeys
on typeFill(t)
	return my typeThen(t, true)
end typeFill

on typeThen(t, fill)
	my toFront()
	-- a line break inside the text is typed as Excel for Mac takes one in a cell, Control+Option+Return
	set AppleScript's text item delimiters to linefeed
	set parts to text items of t
	set AppleScript's text item delimiters to ""
	tell application "System Events"
		repeat with i from 1 to count of parts
			if i > 1 then keystroke return using {control down, option down}
			keystroke (item i of parts)
		end repeat
		delay 0.2
		if fill then
			keystroke return using control down
		else
			keystroke return
		end if
	end tell
	delay 1
	return my answerDialog()
end typeThen

-- Excel brought in front, for up to 3 seconds
on toFront()
	repeat 10 times
		tell application "System Events" to set fp to name of first application process whose frontmost is true
		if fp is "Microsoft Excel" then return
		tell application "Microsoft Excel" to activate
		delay 0.3
	end repeat
	error "Excel is not in front: " & fp
end toFront

-- Excel's windows that are not books, answered in turn: No to a correction it offers, OK to the window that says why a
-- formula is not taken or to an information alert, Cancel to a stop alert of data validation (Retry, Cancel); then
-- Escape leaves the cell as it was. Their texts.
on answerDialog()
	set said to ""
	repeat 5 times
		set found to false
		tell application "System Events" to tell process "Microsoft Excel"
			repeat with w in windows
				set sr to ""
				try
					set sr to subrole of w
				end try
				if sr is not "AXStandardWindow" then
					set tx to ""
					repeat with x in (static texts of w)
						try
							set v to value of x
							if v is not missing value then set tx to tx & (v as text) & " "
						end try
					end repeat
					set said to said & tx & "| "
					set names to name of every button of w
					if names contains "No" then
						click button "No" of w
					else if names contains "Нет" then
						click button "Нет" of w
					else if names contains "OK" then
						click button "OK" of w
					else if names contains "ОК" then
						click button "ОК" of w
					else if names contains "Cancel" then
						click button "Cancel" of w
					else if names contains "Отмена" then
						click button "Отмена" of w
					else
						error "an Excel window without No, OK or Cancel: " & tx
					end if
					set found to true
					exit repeat
				end if
			end repeat
		end tell
		if not found then exit repeat
		delay 1
	end repeat
	if said is "" then return ""
	tell application "System Events" to key code 53
	delay 0.5
	return "dialog: " & said
end answerDialog

-- The formula bar of the window of the book `bookName`: its text area, which the window lists only with Excel in front,
-- its desktop shown: looked for again until it does.
on formulaBar(bookName)
	set bar to missing value
	set seen to ""
	repeat 10 times
		tell application "System Events" to set frontmost of process "Microsoft Excel" to true
		tell application "Microsoft Excel" to activate
		delay 1.5
		try
			tell application "System Events" to tell process "Microsoft Excel"
				set seen to ""
				repeat with w in windows
					set wn to name of w
					set seen to seen & wn & "; "
					if wn is bookName or bookName starts with wn then
						-- a list: a repeat over the reference itself fails
						set els to entire contents of w
						repeat with e in els
							try
								if role of e is "AXTextArea" then
									set bar to contents of e
									exit repeat
								end if
							end try
						end repeat
					end if
					if bar is not missing value then exit repeat
				end repeat
			end tell
		end try
		if bar is not missing value then exit repeat
	end repeat
	if bar is missing value then error "no formula bar in the window of " & bookName & " among " & seen
	return bar
end formulaBar

-- What the formula bar shows: its description ("Formula Bar. A,1 <text>") and its value, read until two reads agree, as
-- the bar follows the selection with a lag.
on barText(bar)
	delay 0.3
	set d to my barNow(bar)
	repeat 10 times
		delay 0.3
		set again to my barNow(bar)
		if again = d then exit repeat
		set d to again
	end repeat
	return d
end barText

on barNow(bar)
	tell application "System Events"
		set d to description of bar
		set v to value of bar
	end tell
	if d is missing value then set d to ""
	if v is missing value then set v to ""
	return {d as text, v as text}
end barNow

-- The element of Excel in focus: its role and value (the cell's editor once its editing began).
on focusedText()
	try
		tell application "System Events" to tell process "Microsoft Excel"
			set f to value of attribute "AXFocusedUIElement"
			set r to role of f
			set v to ""
			try
				set v to value of f
			end try
		end tell
		if v is missing value then set v to ""
		return r & ": " & (v as text)
	on error msg
		return "?: " & msg
	end try
end focusedText
