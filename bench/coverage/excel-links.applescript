-- Open a book with links to other workbooks without updating them, as a person who answers "Don't Update" does; list
-- the linked books (link sources) and the book's folder; read the rectangles (after opening, then after a full
-- calculation); save the book as xlsx under another name and close it.
-- A read takes, for every rectangle, the values of its mirrors on the sheets WT, WN and WX (formulas that never give an
-- error), the rectangle's formulas, its formulas in the user's language and each cell's shown text (`string value`).
-- Arguments: the book's path and the path to save to (both inside Excel's sandbox), the saved file's name, the
-- rectangles as lines "sheet<TAB>range<TAB>mirror range<TAB>its cells, separated by commas".
-- Output: JSON {"sources": [name], "folder": path, "reads": [[[types, numbers, texts, formulas, local formulas, shown], ...], ...]}.
use AppleScript version "2.4"
use framework "Foundation"
use scripting additions

on run argv
	set p to item 1 of argv
	set q to item 2 of argv
	set savedName to item 3 of argv
	set rects to paragraphs of item 4 of argv
	set hfs to (POSIX file p) as text
	set out to (POSIX file q) as text
	set reads to {}
	tell application "Microsoft Excel"
		with timeout of 90 seconds
			set wb to open workbook workbook file name hfs update links do not update links
			set srcs to link sources wb type link type Excel links
			if srcs is missing value then set srcs to {}
			set folderName to path of wb
			repeat with pass from 1 to 2
				if pass is 2 then calculate full
				set snap to {}
				repeat with rl in rects
					set AppleScript's text item delimiters to (character id 9)
					set rp to text items of rl
					set AppleScript's text item delimiters to ""
					set mr to item 3 of rp
					set rg to range (item 2 of rp) of worksheet (item 1 of rp) of wb
					set t to value of range mr of worksheet "WT" of wb
					set nums to value of range mr of worksheet "WN" of wb
					set strs to value of range mr of worksheet "WX" of wb
					set f to formula of rg
					set fl to formula local of rg
					set shown to {}
					set AppleScript's text item delimiters to ","
					set addrs to text items of (item 4 of rp)
					set AppleScript's text item delimiters to ""
					repeat with ad in addrs
						set tx to string value of range (ad as text) of worksheet (item 1 of rp) of wb
						if tx is missing value then set tx to ""
						set end of shown to (tx as text)
					end repeat
					set end of snap to {t, nums, strs, f, fl, shown}
				end repeat
				set end of reads to snap
			end repeat
			save workbook as wb filename out file format Excel XML file format
			close workbook savedName saving no
		end timeout
	end tell
	set d0 to current application's NSDictionary's dictionaryWithObjects:{srcs, folderName, reads} forKeys:{"sources", "folder", "reads"}
	set {d, e} to current application's NSJSONSerialization's dataWithJSONObject:d0 options:0 |error|:(reference)
	if d is missing value then error "JSON: " & (e's localizedDescription() as text)
	return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
