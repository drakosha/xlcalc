-- Copy a range through the object model alone and read plain text without activating Excel.
use framework "Foundation"
use scripting additions
on run argv
  set hfs to (POSIX file (item 1 of argv)) as text
  set js to current application's NSString's stringWithString:(item 3 of argv)
  set acts to (current application's NSJSONSerialization's JSONObjectWithData:(js's dataUsingEncoding:4) options:0 |error|:(missing value)) as list
  set beforeText to the clipboard as text
  set the clipboard to "xlcalc-copy-probe"
  set got to ""
  tell application "Microsoft Excel"
    set wb to open workbook workbook file name hfs
    try
      repeat with act in acts
        set cmd to item 1 of act
        set rg to range (item 2 of act) of worksheet "S" of wb
        if cmd is "hide" then set hidden of entire row of rg to true
        if cmd is "hideColumns" then set hidden of entire column of rg to true
        if cmd is "filter" then autofilter range rg field 1 criteria1 "c"
      end repeat
      calculate full
      copy range (range (item 2 of argv) of worksheet "S" of wb)
      set got to do shell script "/usr/bin/pbpaste | /usr/bin/base64"
    on error msg number n
      close wb saving no
      set the clipboard to beforeText
      error msg number n
    end try
    close wb saving no
  end tell
  set the clipboard to beforeText
  set d to current application's NSJSONSerialization's dataWithJSONObject:{got} options:0 |error|:(missing value)
  return (current application's NSString's alloc()'s initWithData:d encoding:4) as text
end run
