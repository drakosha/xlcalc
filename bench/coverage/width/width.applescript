-- Read fixed and object-model column widths in points, without activating Excel.
use framework "Foundation"
use scripting additions
on run argv
  set hfs to (POSIX file (item 1 of argv)) as text
  set savedPath to (POSIX file (item 2 of argv)) as text
  set spec to current application's NSString's stringWithString:(item 4 of argv)
  set settings to (current application's NSJSONSerialization's JSONObjectWithData:(spec's dataUsingEncoding:4) options:0 |error|:(missing value)) as list
  set widths to {}
  tell application "Microsoft Excel"
    set wb to open workbook workbook file name hfs
    try
      set ws to worksheet "S" of wb
      set baseWidth to standard width of ws
      repeat with pair in settings
        set column width of range ((item 1 of pair) & "1") of ws to item 2 of pair
      end repeat
      repeat with cl in characters of "ABCDEFGHIJKLMNOPQR"
        set rg to range (cl & "1") of ws
        set end of widths to {column width of rg, width of rg, use standard width of rg}
      end repeat
      save workbook as wb filename savedPath file format Excel XML file format
      close workbook (item 3 of argv) saving no
    on error msg number n
      close wb saving no
      error msg number n
    end try
  end tell
  set obj to current application's NSDictionary's dictionaryWithObjects:{baseWidth, widths} forKeys:{"standardWidth", "columns"}
  set d to current application's NSJSONSerialization's dataWithJSONObject:obj options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
