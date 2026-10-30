-- A book's cells read as it opens, after a recalculation, after column A of S and the sheet T are deleted, and after a
-- full calculation; the book saved as xlsx after the deletion. Each read: [formula, shown text] of S's cells listed.
use framework "Foundation"
use scripting additions
on readCells(ws, addrs)
  set got to {}
  tell application "Microsoft Excel"
    repeat with ad in addrs
      set rg to range (ad as text) of ws
      set end of got to {formula of rg, string value of rg}
    end repeat
  end tell
  return got
end readCells
on run argv
  set hfs to (POSIX file (item 1 of argv)) as text
  set savedPath to (POSIX file (item 2 of argv)) as text
  set spec to current application's NSString's stringWithString:(item 4 of argv)
  set addrPair to (current application's NSJSONSerialization's JSONObjectWithData:(spec's dataUsingEncoding:4) options:0 |error|:(missing value)) as list
  set cellsBefore to item 1 of addrPair
  set cellsAfter to item 2 of addrPair
  tell application "Microsoft Excel"
    set wb to open workbook workbook file name hfs
    try
      set ws to worksheet "S" of wb
      set opened to my readCells(ws, cellsBefore)
      calculate
      set recalculated to my readCells(ws, cellsBefore)
      set display alerts to false
      delete range (entire column of (range "A1" of ws)) shift shift to left
      delete worksheet "T" of wb
      set display alerts to true
      set deleted to my readCells(ws, cellsAfter)
      save workbook as wb filename savedPath file format Excel XML file format
      set ws to worksheet "S" of active workbook
      calculate full
      set fullRead to my readCells(ws, cellsAfter)
      close workbook (item 3 of argv) saving no
    on error msg number n
      set display alerts to true
      close wb saving no
      error msg number n
    end try
  end tell
  set obj to current application's NSDictionary's dictionaryWithObjects:{opened, recalculated, deleted, fullRead} forKeys:{"opened", "recalculated", "deleted", "full"}
  set d to current application's NSJSONSerialization's dataWithJSONObject:obj options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
