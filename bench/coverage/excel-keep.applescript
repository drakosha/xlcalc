-- Open a book and give it what the workbook model does not hold, as a person does in Excel's window, then save it as
-- xlsx under another name and close it. Arguments: the book's path and the path to save to (inside Excel's sandbox),
-- the saved file's name, the features (a comma list: chart, chartsheet, picture, sized, floating, chartmove, rich, theme,
-- print, breaks, note, autofilter, criteria, overtable, tableover, tablepart, tablein, tablearound) and
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
				set pic to make new picture at ws with properties {file name:img, link to file:false, save with document:true, top:220, left position:300, width:80, height:50}
				-- moving and sizing with its cells, or fixed where it is, as the features ask; Excel's own is moving only
				if feats contains "sized" then set placement of pic to placement move and size
				set made to made & "picture" & linefeed
			end if
			if feats contains "floating" then
				set pic to make new picture at ws with properties {file name:img, link to file:false, save with document:true, top:300, left position:120, width:60, height:40}
				set placement of pic to placement free floating
				set made to made & "floating" & linefeed
			end if
			if feats contains "chartmove" then
				-- the chart moving with its cells and keeping its size
				set placement of (chart object 1 of ws) to placement move
				set made to made & "chartmove" & linefeed
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
			if feats contains "autofilter" then
				-- an AutoFilter over the amounts without criteria, as Data > Filter puts it
				autofilter range (range "A1:C21" of ws) field 1
				set made to made & "autofilter " & (get address of (range object of (autofilter object of ws))) & linefeed
			end if
			if feats contains "criteria" then
				autofilter range (range "A1:C21" of ws) field 1 criteria1 "North"
				set made to made & "criteria" & linefeed
			end if
			if feats contains "overtable" then
				-- an AutoFilter asked over Plain!A1:B5, where a table lies on A1:A2
				set wp to worksheet "Plain" of wb
				try
					autofilter range (range "A1:B5" of wp) field 1
					set made to made & "overtable done; autofilter mode " & (autofilter mode of wp) & linefeed
				on error msg number k
					set made to made & "overtable error " & k & " " & msg & linefeed
				end try
			end if
			if feats contains "tableover" then
				-- a table asked over the AutoFilter's range
				try
					make new list object at ws with properties {range object:(range "A1:C21" of ws)}
					set made to made & "tableover done; autofilter mode " & (autofilter mode of ws) & "; tables " & (count of list objects of ws) & linefeed
				on error msg number k
					set made to made & "tableover error " & k & " " & msg & linefeed
				end try
			end if
			-- a table asked over part of the AutoFilter's data and the column beside it, inside its range, around it
			repeat with probe in {{"tablepart", "B10:D12"}, {"tablein", "A1:B10"}, {"tablearound", "A1:D22"}}
				if feats contains (item 1 of probe) then
					try
						make new list object at ws with properties {range object:(range (item 2 of probe) of ws)}
						set made to made & (item 1 of probe) & " done; autofilter mode " & (autofilter mode of ws) & "; tables " & (count of list objects of ws) & linefeed
					on error msg number k
						set made to made & (item 1 of probe) & " error " & k & " " & msg & linefeed
					end try
				end if
			end repeat
			calculate full
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	return made
end run
