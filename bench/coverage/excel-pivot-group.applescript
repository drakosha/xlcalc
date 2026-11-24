-- Group a PivotTable row field in the Russian Grouping dialog, save the book and return TableRange1.
on run argv
	set sourcePath to (POSIX file (item 1 of argv)) as text
	set savedPath to (POSIX file (item 2 of argv)) as text
	set savedName to item 3 of argv
	set flags to item 4 of argv
	set lowerBound to item 5 of argv
	set upperBound to item 6 of argv
	set groupStep to item 7 of argv
	tell application "Microsoft Excel"
		with timeout of 170 seconds
			set wb to open workbook workbook file name sourcePath update links do not update links
			try
				set ws to worksheet "P" of wb
				activate object ws
				activate
				-- The input has one tabular row field and its header is in A3.
				select range "A4" of ws
				my groupSelection(flags, lowerBound, upperBound, groupStep)
				set pt to pivot table "PT1" of ws
				set areaAddress to get address of table range1 of pt
				save workbook as wb filename savedPath file format Excel XML file format
			on error msg number n
				try
					close wb saving no
				end try
				error msg number n
			end try
			close workbook savedName saving no
			return areaAddress
		end timeout
	end tell
end run

on groupSelection(flags, lowerBound, upperBound, groupStep)
	tell application "System Events" to tell process "Microsoft Excel"
		set dataMenu to menu 1 of menu bar item "Данные" of menu bar 1
		set groupMenu to menu 1 of menu item "Группа и структура" of dataMenu
		click menu item "Сгруппировать..." of groupMenu
		set groupingDialog to window "Группировка"
		repeat 20 times
			if exists groupingDialog then exit repeat
			delay 0.1
		end repeat
		if not (exists groupingDialog) then error "Excel did not show the Grouping dialog"
		try
			set inputs to text fields of groupingDialog
			if lowerBound is not "" then my setBound(groupingDialog, item 1 of inputs, lowerBound)
			if upperBound is not "" then my setBound(groupingDialog, item 2 of inputs, upperBound)
			if flags is "" then
				set value of item 3 of inputs to groupStep
			else
				set periodTable to my dialogElement(groupingDialog, "AXTable", missing value)
				set periodRows to rows of periodTable
				if (count periodRows) is not (length of flags) then error "Excel did not show the seven grouping periods"
				set chosenRows to {}
				repeat with k from 1 to count periodRows
					if character k of flags is "1" then set end of chosenRows to item k of periodRows
				end repeat
				my clickPeriods(chosenRows, false)
				delay 0.2
				set toggles to {}
				repeat with k from 1 to count periodRows
					set periodRow to item k of periodRows
					set isChosen to value of attribute "AXSelected" of periodRow
					if isChosen is not (character k of flags is "1") then set end of toggles to periodRow
				end repeat
				if (count toggles) is greater than 0 then my clickPeriods(toggles, true)
				if groupStep is not "1" then
					set daysInput to my dialogElement(groupingDialog, "AXTextField", "Число дней:")
					set value of attribute "AXFocused" of daysInput to true
					keystroke "a" using command down
					keystroke groupStep
					key code 48
				end if
			end if
			click button "ОК" of groupingDialog
			repeat 20 times
				if not (exists groupingDialog) then return
				delay 0.1
			end repeat
			error "Excel did not apply the grouping"
		on error msg number n
			if exists groupingDialog then click button "Отмена" of groupingDialog
			error msg number n
		end try
	end tell
end groupSelection

on setBound(groupingDialog, fieldInput, newValue)
	tell application "System Events"
		set toggle to checkbox (name of fieldInput) of groupingDialog
		if value of toggle is 1 then click toggle
		set value of fieldInput to newValue
	end tell
end setBound

on dialogElement(groupingDialog, roleName, elementName)
	tell application "System Events"
		set elements to entire contents of groupingDialog
		repeat with elementRef in elements
			set el to contents of elementRef
			if role of el is roleName and (elementName is missing value or name of el is elementName) then return el
		end repeat
	end tell
	error "Excel's Grouping dialog has no " & roleName
end dialogElement

-- Setting AXSelectedRows changes the accessibility value without changing Excel's grouping.
on clickPeriods(chosenRows, keepSelection)
	set pointList to {}
	tell application "System Events"
		repeat with rowRef in chosenRows
			set periodLabel to text field 1 of (contents of rowRef)
			set {x, y} to position of periodLabel
			set {w, h} to size of periodLabel
			set end of pointList to "[" & (x + w div 2) & "," & (y + h div 2) & "]"
		end repeat
	end tell
	set AppleScript's text item delimiters to ","
	set commandFirst to "false"
	if keepSelection then set commandFirst to "true"
	set js to "ObjC.import('CoreGraphics');
const commandMask = 1 << 20;
const points = [" & (pointList as text) & "];
points.forEach(function (point, index) {
  const position = $.CGPointMake(point[0], point[1]);
  for (const kind of [1, 2]) {
    const event = $.CGEventCreateMouseEvent(null, kind, position, 0);
    if (" & commandFirst & " || index > 0) $.CGEventSetFlags(event, commandMask);
    $.CGEventPost(0, event);
    delay(0.05);
  }
});"
	do shell script "osascript -l JavaScript -e " & quoted form of js
end clickPeriods
