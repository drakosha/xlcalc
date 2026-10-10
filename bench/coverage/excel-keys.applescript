-- Typing on the keyboard, for the steps that type texts as a user does (excel-entry.applescript, excel-fit.applescript),
-- which load these handlers with `keysLib`: no `use` statements here, as the loader wraps the file in a script object.
-- a text typed into the selected cell, then Return; Excel must be the application in front (it is brought forward for up
-- to 3 seconds). A window Excel shows for it (a formula it does not take) is answered, and Escape leaves the cell as it
-- was: its text is returned.
on typeKeys(t)
	repeat 10 times
		tell application "System Events" to set fp to name of first application process whose frontmost is true
		if fp is "Microsoft Excel" then exit repeat
		tell application "Microsoft Excel" to activate
		delay 0.3
	end repeat
	tell application "System Events"
		set fp to name of first application process whose frontmost is true
		if fp is not "Microsoft Excel" then error "Excel is not in front: " & fp
		keystroke t
		delay 0.2
		keystroke return
	end tell
	delay 1
	return my answerDialog()
end typeKeys

-- Excel's windows that are not books, answered in turn: No to a correction it offers, OK to the window that says why a
-- formula is not taken; then Escape leaves the cell as it was. Their texts.
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
					else
						error "an Excel window without No or OK: " & tx
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
