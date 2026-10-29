-- Open a rows part's book, read each case's row height, and case by case (a row of the sheet "S") type its texts in
-- turn, set a font size or a rotation (`text orientation`) on its columns whole when it patches them, and fit the row
-- (`autofit` of its entire row) when it has no texts and no patch or the case says so: texts through `value` or
-- `formula2 local` with alerts off, or with Excel in front (System Events): typed into the cell and Return (`keys`), into
-- a selection and Control+Return (`fill`), a line break typed as Control+Option+Return, or put on the clipboard as plain
-- text and pasted with Cmd+V (`paste`); the windows Excel shows for them are answered. Then read each case's row: `row
-- height`, `use standard height`, its cells' `formula` and `wrap text` and the heights of the other rows it names; save
-- the book as xlsx under another name and close it.
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the cases as
-- a JSON array of [row, fit (0 or 1), [[cell, text, via, last cell of a fill's selection], ...], [cells to read],
-- [columns patched, font size or 0, rotation: degrees, "vertical" or ""] or [], [rows read after]], the path of
-- excel-keys.applescript.
-- Output: JSON {"before": [row height, ...], "refused": [[Excel's error for each text, "" when it took it], ...],
-- "rows": [[row height, standard, [[formula, wrap], ...], [row height of each row read after]], ...]}.
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
					else if via is "fill" then
						activate
						delay 0.5
						select range ((item 1 of x) & ":" & (item 4 of x)) of ws
						set er to kb's typeFill(t)
					else if via is "paste" then
						my toClipboard(t)
						activate
						delay 0.5
						select rg
						kb's toFront()
						-- the V key, whatever the layout types on it, as a person pastes
						tell application "System Events" to key code 9 using command down
						delay 1.5
						set er to kb's answerDialog()
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
				set pt to item 5 of c
				if (count of pt) is 3 then
					set cols to range ((item 1 of pt) as text) of ws
					if (item 2 of pt) as integer > 0 then set font size of font object of cols to (item 2 of pt) as integer
					if (item 3 of pt) as text is "vertical" then
						set text orientation of cols to orientation vertical
					else if (item 3 of pt) as text is not "" then
						set text orientation of cols to ((item 3 of pt) as integer)
					end if
				end if
				if (item 2 of c) is 1 then autofit rw
			end repeat
			calculate
			repeat with c in spec
				set rw to entire row of range ("A" & ((item 1 of c) as integer)) of ws
				set wraps to {}
				repeat with a in (item 4 of c)
					set cl to range (a as text) of ws
					set end of wraps to {formula of cl, wrap text of cl}
				end repeat
				set others to {}
				repeat with r in (item 6 of c)
					set end of others to row height of (entire row of range ("A" & (r as integer)) of ws)
				end repeat
				set end of rowsOut to {row height of rw, use standard height of rw, wraps, others}
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

-- the text on the general pasteboard as plain text, nothing else with it
on toClipboard(t)
	set pb to current application's NSPasteboard's generalPasteboard()
	pb's clearContents()
	pb's setString:t forType:(current application's NSPasteboardTypeString)
end toClipboard

-- the handlers of excel-keys.applescript (its path `p`) as a script object
on keysLib(p)
	set src to (current application's NSString's stringWithContentsOfFile:p encoding:4 |error|:(missing value)) as text
	return run script ("script K" & linefeed & src & linefeed & "end script" & linefeed & "return K")
end keysLib
