-- Open a book, take the steps (a number format, bold or a fill set on a range, a value typed into a cell), compute the
-- book, read what the listed cells show (Range.Text, the number format, the display format's fill and font colours, bold
-- and italic), save the book as xlsx under another name and close it. A step Excel refuses comes back as `error`.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the steps as
-- JSON ([[sheet, range, number format or "", bold, [r, g, b] of a fill or [], value kind "n", "s" or "", value]]), the
-- cells to read as JSON ([[sheet, range]]).
-- Output: JSON {"error": text, "looks": [[sheet, [[address, text, number format, [r, g, b] of the fill, its colour index,
-- [r, g, b] of the font, bold, italic]]]]}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on parsed(t)
	set d to (current application's NSString's stringWithString:t)'s dataUsingEncoding:4
	return (current application's NSJSONSerialization's JSONObjectWithData:d options:0 |error|:(missing value)) as list
end parsed

on txt(v)
	if v is missing value then return ""
	return v as text
end txt

-- a typed number through NSString, as `as number` reads the system's decimal separator
on num(s)
	return ((current application's NSString's stringWithString:s)'s doubleValue()) as real
end num

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set steps to parsed(item 4 of argv)
	set reads to parsed(item 5 of argv)
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set stepError to ""
	set readError to ""
	set looksOut to {}
	tell application "Microsoft Excel"
		with timeout of 120 seconds
			set wb to open workbook workbook file name hfs
			repeat with st in steps
				try
					set ws to worksheet ((item 1 of st) as text) of wb
					set rg to range ((item 2 of st) as text) of ws
					if (item 3 of st) as text is not "" then set number format of rg to ((item 3 of st) as text)
					if (item 4 of st) is true then set bold of font object of rg to true
					if (count of (item 5 of st)) is 3 then set color of interior object of rg to (item 5 of st)
					if (item 6 of st) as text is "n" then
						set value of rg to (my num((item 7 of st) as text))
					else if (item 6 of st) as text is "s" then
						set value of rg to ((item 7 of st) as text)
					end if
				on error msg number k
					set stepError to stepError & (item 2 of st) & ": " & (k as text) & " " & msg & "; "
				end try
			end repeat
			calculate full
			try
				repeat with readItem in reads
					set ws to worksheet ((item 1 of readItem) as text) of wb
					-- `cell j of` a range of a sheet that is not active gives the active sheet's cell
					activate object ws
					set rng to range ((item 2 of readItem) as text) of ws
					set cellsOut to {}
					repeat with j from 1 to (count of cells of rng)
						set oneCell to cell j of rng
						set dispFmt to display format of oneCell
						set intObj to interior object of dispFmt
						set fontObj to font object of dispFmt
						set end of cellsOut to {(get address oneCell), my txt(string value of oneCell), my txt(number format of oneCell), (color of intObj), ((color index of intObj) as text), (color of fontObj), (bold of fontObj), (italic of fontObj)}
					end repeat
					set end of looksOut to {name of ws, cellsOut}
				end repeat
			on error msg number k
				set readError to (k as text) & " " & msg
			end try
			if readError is "" then
				activate object (worksheet 1 of wb)
				save workbook as wb filename out file format Excel XML file format
				close workbook savedName saving no
			else
				close wb saving no
			end if
		end timeout
	end tell
	if readError is not "" then error "reading: " & readError
	set d0 to current application's NSDictionary's dictionaryWithObjects:{stepError, looksOut} forKeys:{"error", "looks"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
