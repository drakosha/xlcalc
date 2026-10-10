-- Open a fit part's book and, case by case (a column of the sheet "S"), fit the column (`autofit` of its entire column)
-- or type its texts in turn, after fitting it when the case says so: through `value` or `formula2 local` with alerts off,
-- or on the keyboard (System Events, Excel in front, checked before each text, each followed by Return; the window
-- Excel shows for a formula it does not take is answered and its text is the text's refusal). Then read each case's
-- column: `column width`, `width` (points), `use standard width`, and what its rows show (`string value`); save the
-- book as xlsx under another name and close it.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the cases as
-- a JSON array of [column letter, fit first (0 or 1), [[row, text, via], ...], [rows to read]] (a case without texts
-- is fitted), the path of excel-keys.applescript.
-- Output: JSON {"refused": [[Excel's error for each text, "" when it took it], ...], "cases": [[column width, points,
-- standard, [texts]], ...]}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set js to current application's NSString's stringWithString:(item 4 of argv)
	set spec to (current application's NSJSONSerialization's JSONObjectWithData:(js's dataUsingEncoding:4) options:0 |error|:(missing value)) as list
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set refusedOut to {}
	set casesOut to {}
	set kb to my keysLib(item 5 of argv)
	tell application "Microsoft Excel"
		with timeout of 280 seconds
			set wb to open workbook workbook file name hfs
			set ws to worksheet "S" of wb
			repeat with c in spec
				set cl to item 1 of c
				set texts to item 3 of c
				set ers to {}
				if (count of texts) is 0 or (item 2 of c) is 1 then autofit (entire column of range (cl & "1") of ws)
				repeat with x in texts
					set rg to range (cl & ((item 1 of x) as integer)) of ws
					set t to item 2 of x
					set via to item 3 of x
					set er to ""
					if via is "keys" then
						activate
						delay 0.5
						select rg
						set er to kb's typeKeys(t)
					else
						set display alerts to false
						try
							if via is "value" then
								set value of rg to t
							else
								set formula2 local of rg to t
							end if
						on error msg number k
							set er to (k as text) & " " & msg
						end try
						set display alerts to true
					end if
					set end of ers to er
				end repeat
				set end of refusedOut to ers
			end repeat
			calculate
			repeat with c in spec
				set cl to item 1 of c
				set firstCell to range (cl & "1") of ws
				set shown to {}
				repeat with r in (item 4 of c)
					set tx to string value of range (cl & (r as integer)) of ws
					if tx is missing value then set tx to ""
					set end of shown to tx as text
				end repeat
				set end of casesOut to {column width of firstCell, width of firstCell, use standard width of firstCell, shown}
			end repeat
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{refusedOut, casesOut} forKeys:{"refused", "cases"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run

-- the handlers of excel-keys.applescript (its path `p`) as a script object
on keysLib(p)
	set src to (current application's NSString's stringWithContentsOfFile:p encoding:4 |error|:(missing value)) as text
	return run script ("script K" & linefeed & src & linefeed & "end script" & linefeed & "return K")
end keysLib
