-- Open a probe book in Excel, read column A of the case sheet C as formula text twice, in English
-- (`formula`) and in the language of the user (`formula local`), and close it without saving. Nothing is
-- calculated or changed. Output: JSON [[formula], [formula local]].
-- Arguments: the book's path (inside Excel's sandbox), the number of rows.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set n to item 2 of argv
	set hfs to (POSIX file p) as text
	tell application "Microsoft Excel"
		with timeout of 150 seconds
			set wb to open workbook workbook file name hfs
			set ws to worksheet "C" of wb
			set en to formula of range ("A1:A" & n) of ws
			set loc to formula local of range ("A1:A" & n) of ws
			close wb saving no
		end timeout
	end tell
	set arr to current application's NSArray's arrayWithArray:{en, loc}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:arr options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
