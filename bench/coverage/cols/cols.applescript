-- Read each listed column's width and its first cell's number format, then save the book as xlsx.
use framework "Foundation"
use scripting additions
on run argv
  set hfs to (POSIX file (item 1 of argv)) as text
  set savedPath to (POSIX file (item 2 of argv)) as text
  set spec to current application's NSString's stringWithString:(item 4 of argv)
  set letters to (current application's NSJSONSerialization's JSONObjectWithData:(spec's dataUsingEncoding:4) options:0 |error|:(missing value)) as list
  set got to {}
  tell application "Microsoft Excel"
    set wb to open workbook workbook file name hfs
    try
      set ws to worksheet "S" of wb
      repeat with cl in letters
        set rg to range ((cl as text) & "1") of ws
        set end of got to {column width of rg, number format of rg}
      end repeat
      save workbook as wb filename savedPath file format Excel XML file format
      close workbook (item 3 of argv) saving no
    on error msg number n
      close wb saving no
      error msg number n
    end try
  end tell
  set d to current application's NSJSONSerialization's dataWithJSONObject:got options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
