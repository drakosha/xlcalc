-- Open a book and give it what the workbook model does not hold, as a person does in Excel's window, then save it as
-- xlsx under another name and close it. Arguments: the book's path and the path to save to (inside Excel's sandbox),
-- the saved file's name, the features (a comma list: chart, chartsheet, picture, rich, theme, print, breaks, note) and
-- a picture's path (inside the sandbox). Answers what it made, one feature a line, or the error of one it could not.
on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set feats to item 4 of argv
	set img to item 5 of argv
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set made to ""
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			set wb to open workbook workbook file name hfs
			set ws to worksheet "Data" of wb
			if feats contains "chart" then
				-- a column chart over the amounts, beside them
				set co to make new chart object at ws with properties {left position:300, top:10, width:300, height:180}
				set ch to chart of co
				set chart type of ch to column clustered
				set source data ch source (range "A1:B6" of ws)
				set made to made & "chart" & linefeed
			end if
			if feats contains "chartsheet" then
				-- a line chart of the months on a sheet of its own
				set co to make new chart object at ws with properties {left position:300, top:200, width:300, height:180}
				set ch to chart of co
				set chart type of ch to line markers
				set source data ch source (range "B1:B21" of ws)
				chart location ch where location as new sheet name "Trend"
				set made to made & "chartsheet" & linefeed
			end if
			if feats contains "picture" then
				make new picture at ws with properties {file name:img, link to file:false, save with document:true, top:220, left position:300, width:80, height:50}
				set made to made & "picture" & linefeed
			end if
			if feats contains "rich" then
				-- formatted runs in a cell's text
				set bold of font object of characters 1 thru 3 of range "A2" of ws to true
				set italic of font object of characters 2 thru 4 of range "A5" of ws to true
				set made to made & "rich" & linefeed
			end if
			if feats contains "theme" then
				set theme color of font object of range "A3:C3" of ws to first accent theme color
				set theme color of interior object of range "A4:C4" of ws to second accent theme color
				set made to made & "theme" & linefeed
			end if
			if feats contains "print" then
				set print area of page setup object of ws to "$A$1:$C$21"
				set made to made & "print" & linefeed
			end if
			if feats contains "breaks" then
				make new horizontal page break at ws with properties {location:range "A12" of ws}
				set made to made & "breaks" & linefeed
			end if
			if feats contains "note" then
				add comment (range "C2" of ws) comment text "made in Excel"
				set made to made & "note" & linefeed
			end if
			calculate full
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	return made
end run
