-- Open a rows part's book, read each case's row height, and case by case (a row of the sheet "S") type its texts in
-- turn and fit the row (`autofit` of its entire row) when it has no texts or the case says so: texts through `value` or
-- `formula2 local` with alerts off, or on the keyboard (System Events, Excel in front, a line break typed as
-- Control+Option+Return, each text followed by Return). Then read each case's row: `row height`, `use standard
-- height`, and whether its cells wrap (`wrap text`); save the book as xlsx under another name and close it.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the cases as
-- a JSON array of [row, fit (0 or 1), [[cell, text, via], ...], [cells to read]], the path of excel-keys.applescript.
-- Output: JSON {"before": [row height, ...], "refused": [[Excel's error for each text, "" when it took it], ...],
-- "rows": [[row height, standard, [wraps]], ...]}.
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
	set beforeOut to {}
	set refusedOut to {}
	set rowsOut to {}
	set kb to my keysLib(item 5 of argv)
	tell application "Microsoft Excel"
		with timeout of 280 seconds
			set wb to open workbook workbook file name hfs
			set ws to worksheet "S" of wb
			repeat with c in spec
				set end of beforeOut to row height of (entire row of range ("A" & ((item 1 of c) as integer)) of ws)
			end repeat
			repeat with c in spec
				set rw to entire row of range ("A" & ((item 1 of c) as integer)) of ws
				set texts to item 3 of c
				set ers to {}
				repeat with x in texts
					set rg to range (item 1 of x) of ws
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
				if (item 2 of c) is 1 then autofit rw
			end repeat
			calculate
			repeat with c in spec
				set rw to entire row of range ("A" & ((item 1 of c) as integer)) of ws
				set wraps to {}
				repeat with a in (item 4 of c)
					set end of wraps to wrap text of range (a as text) of ws
				end repeat
				set end of rowsOut to {row height of rw, use standard height of rw, wraps}
			end repeat
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{beforeOut, refusedOut, rowsOut} forKeys:{"before", "refused", "rows"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run

-- the handlers of excel-keys.applescript (its path `p`) as a script object
on keysLib(p)
	set src to (current application's NSString's stringWithContentsOfFile:p encoding:4 |error|:(missing value)) as text
	return run script ("script K" & linefeed & src & linefeed & "end script" & linefeed & "return K")
end keysLib
