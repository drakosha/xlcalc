-- Open a file (a CSV file, or a book), save it as xlsx under another name, then as CSV (comma delimited, Excel's CSV
-- file format) and close it. Alerts are off while it saves: saving a book as CSV keeps its active sheet alone.
-- Arguments: the file's path, the path to save the xlsx to and its name, the path to save the CSV to and its name
-- (all inside Excel's sandbox). Output: the name of the active sheet as Excel opened the file.
on run argv
	set hfs to (POSIX file (item 1 of argv)) as text
	set out to (POSIX file (item 2 of argv)) as text
	set savedName to item 3 of argv
	set csvOut to (POSIX file (item 4 of argv)) as text
	set csvName to item 5 of argv
	tell application "Microsoft Excel"
		with timeout of 90 seconds
			set wb to open workbook workbook file name hfs update links do not update links
			set sheetName to name of active sheet of wb
			set display alerts to false
			save workbook as wb filename out file format Excel XML file format
			save workbook as workbook savedName filename csvOut file format CSV file format
			close workbook csvName saving no
			set display alerts to true
		end timeout
	end tell
	return sheetName
end run
