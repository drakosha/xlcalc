# Runs one job of the Windows Excel harness (bench/windows): opens each book of <dir>\job.json in one Excel, takes
# the book's steps and writes its answer to <dir>\out\<n>.json as soon as it is done, so a job cut short keeps the
# books done. Excel is quit at the end with no book left open.
# A step is { do: ..., ... } and answers one JSON value: null for an action, the cells read, or { "error": text }
# when Excel refused the step. Cells: null blank, a number, a string, a boolean, { "e": code } an error (COM's code).
param([Parameter(Mandatory = $true)][string]$dir)
$ErrorActionPreference = 'Stop'
[Console]::OutputEncoding = [Text.Encoding]::UTF8
$inv =[Globalization.CultureInfo]::InvariantCulture
$job = Get-Content -Raw -Encoding UTF8 (Join-Path $dir 'job.json') | ConvertFrom-Json
$out = Join-Path $dir 'out'
New-Item -ItemType Directory -Force $out | Out-Null
# the outline step reads the copies Excel saves
Add-Type -AssemblyName System.IO.Compression.FileSystem

# JSON text: everything outside printable ASCII escaped, so the file is ASCII and lone surrogates survive
function Str([string]$s) { '"' + [regex]::Replace($s, '[^\x20\x21\x23-\x5b\x5d-\x7e]', { param($m) '\u{0:x4}' -f [int][char]$m.Value }) + '"' }

function Cell($v) {
  if ($null -eq $v) { 'null' }
  elseif ($v -is [double]) { $v.ToString('G17', $inv) }
  elseif ($v -is [string]) { Str $v }
  elseif ($v -is [bool]) { if ($v) { 'true' } else { 'false' } }
  elseif ($v -is [int]) { '{"e":' + $v + '}' }
  else { '{"other":' + (Str ("$($v.GetType().Name) $v")) + '}' }
}

# a range's Value2, Formula or Formula2 as rows of cells (one cell comes as a scalar)
function Grid($x) {
  if ($x -isnot [array]) { return '[[' + (Cell $x) + ']]' }
  $rows = for ($r = $x.GetLowerBound(0); $r -le $x.GetUpperBound(0); $r++) {
    $cells = for ($c = $x.GetLowerBound(1); $c -le $x.GetUpperBound(1); $c++) { Cell $x[$r, $c] }
    '[' + ($cells -join ',') + ']'
  }
  '[' + ($rows -join ',') + ']'
}

function List($xs) { '[' + (($xs | ForEach-Object { $_ }) -join ',') + ']' }
function Err($e) { '{"error":' + (Str ("$($e.Exception.HResult) $($e.Exception.Message)")) + '}' }

# a COM range is enumerable: the comma keeps PowerShell from unrolling it into its cells
function Rg($wb, $sheet, $addr) { , $wb.Worksheets.Item($sheet).Range($addr) }

# conditional formats: CondMake, CondLooks, CondRules (bench/windows/cond.ts)
. (Join-Path $PSScriptRoot 'cond.ps1')

# what a cell shows (Range.Text), the colour of its display format's font (BGR) and, with codes, its format codes
function Shown($c, $codes) {
  $t = $c.Text
  $parts = @((Str ([string]$t)), (Cell $c.DisplayFormat.Font.Color))
  if ($codes) { $parts += @((Str $c.NumberFormat), (Str $c.NumberFormatLocal)) }
  List $parts
}

# a structure operation as excel-struct.applescript takes it; Excel's refusal is the step's error
function Struct($wb, $s) {
  $ws = if ($s.sheet) { $wb.Worksheets.Item($s.sheet) } else { $null }
  switch ($s.op) {
    'insertRows' { [void]$ws.Range($s.range).EntireRow.Insert(-4121) }
    'deleteRows' { [void]$ws.Range($s.range).EntireRow.Delete(-4162) }
    'insertColumns' { [void]$ws.Range($s.range).EntireColumn.Insert(-4161) }
    'deleteColumns' { [void]$ws.Range($s.range).EntireColumn.Delete(-4159) }
    'moveRange' { [void]$ws.Range($s.range).Cut((Rg $wb $s.target $s.to)) }
    'copyRange' { [void]$ws.Range($s.range).Copy((Rg $wb $s.target $s.to)) }
    'fillDown' { [void](Rg $wb $s.target $s.to).FillDown() }
    'autofill' { [void]$ws.Range($s.range).AutoFill((Rg $wb $s.target $s.to), 0) }
    # flags in target: across, center (alerts are off, so Excel keeps the top-left value without asking)
    'mergeCells' {
      $rg = $ws.Range($s.range)
      [void]$rg.Merge([bool]($s.target -match 'across'))
      if ($s.target -match 'center') { $rg.HorizontalAlignment = -4108 }
    }
    'unmergeCells' { [void]$ws.Range($s.range).UnMerge() }
    # the fill handle over the range `to`, which holds the source; the fill's type (XlAutoFillType) in target
    'fillRange' { [void]$ws.Range($s.range).AutoFill($ws.Range($s.to), [int]$s.target) }
    # a table with headers over the range, named target (Excel for Windows only; the tables step makes its own)
    'addTable' { $ws.ListObjects.Add(1, $ws.Range($s.range), [Type]::Missing, 1).Name = $s.target }
    'renameSheet' { $ws.Name = $s.target }
    'deleteSheet' { [void]$ws.Delete() }
    'insertSheet' {
      $n = $wb.Worksheets.Count
      $nw = if ($s.sheet) { $wb.Worksheets.Add($wb.Worksheets.Item($s.sheet)) } else { $wb.Worksheets.Add([Type]::Missing, $wb.Worksheets.Item($n)) }
      $nw.Name = $s.target
    }
    'moveSheet' {
      if ($s.target) { $ws.Move($wb.Worksheets.Item($s.target)) } else { $ws.Move([Type]::Missing, $wb.Worksheets.Item($wb.Worksheets.Count)) }
    }
    default { throw "unknown op $($s.op)" }
  }
}

# the look of a sheet: cells [address, bold, fill BGR, format, merged area, horizontal alignment], columns [letter, standard, width, hidden], rows likewise
function Look($ws, $s) {
  $cells = foreach ($a in $s.cells) {
    $c = $ws.Range($a)
    $m = if ($c.MergeCells) { $c.MergeArea.Address() } else { '' }
    List @((Str $a), (Cell $c.Font.Bold), (Cell $c.Interior.Color), (Str $c.NumberFormat), (Str $m), (Cell $c.HorizontalAlignment))
  }
  $cols = foreach ($k in $s.columns) {
    $c = $ws.Range("${k}1")
    List @((Str $k), (Cell $c.UseStandardWidth), (Cell $c.ColumnWidth), (Cell $c.EntireColumn.Hidden))
  }
  $rows = foreach ($r in $s.rows) {
    $c = $ws.Range("A$r")
    List @((Str ([string]$r)), (Cell $c.UseStandardHeight), (Cell $c.RowHeight), (Cell $c.EntireRow.Hidden))
  }
  List @((List $cells), (List $cols), (List $rows))
}

# a step of a table case, a line as excel-tables.applescript takes it (bench/coverage/tables.ts); a formula typed in
# the user's language goes in as Excel 365 enters typed text (Formula2Local)
function Table($wb, [string]$line) {
  $p = $line -split "`t"
  # a name's step has its scope there, empty for a global name
  $ws = if ($p[1]) { $wb.Worksheets.Item($p[1]) } else { $null }
  switch ($p[0]) {
    'e' {
      $rg = $ws.Range($p[2])
      switch ($p[3]) {
        'x' { [void]$rg.ClearContents() }
        # a number typed: its text with a decimal point, as Formula reads it in any locale (Value2 with a Double into a
        # table's header failed now and then with an invalid cast)
        'n' { $rg.Formula = $p[4] }
        'f' { $rg.Formula2 = $p[4] }
        'l' { $rg.Formula2Local = $p[4] }
        default { $rg.Value2 = $p[4] }
      }
    }
    'insr' { [void]$ws.Range($p[2]).EntireRow.Insert(-4121) }
    'delr' { [void]$ws.Range($p[2]).EntireRow.Delete(-4162) }
    'insc' { [void]$ws.Range($p[2]).EntireColumn.Insert(-4161) }
    'delc' { [void]$ws.Range($p[2]).EntireColumn.Delete(-4159) }
    'rename' { $ws.ListObjects.Item($p[2]).Name = $p[3] }
    'deltable' { [void]$ws.ListObjects.Item($p[2]).Delete() }
    'convert' { [void]$ws.ListObjects.Item($p[2]).Unlist() }
    'copy' { [void]$ws.Range($p[2]).Copy((Rg $wb $p[3] $p[4])) }
    'cut' { [void]$ws.Range($p[2]).Cut((Rg $wb $p[3] $p[4])) }
    'renamesheet' { $ws.Name = $p[2] }
    'delsheet' { [void]$ws.Delete() }
    # Names.Add reads its formula in the user's language (=0.2 is an error in ru-RU), RefersTo in English
    'name' { $n = if ($ws) { $ws.Names.Add($p[2], '=0') } else { $wb.Names.Add($p[2], '=0') }; $n.RefersTo = $p[3] }
    default { throw "unknown table step $($p[0])" }
  }
}

# a sheet's notes, threads of comments, hyperlinks and window: [name, [[address, author, text]], [[address, [[author,
# text]]]], [[range, Address, SubAddress, ScreenTip, TextToDisplay]], [frozen, split row, split column, the first row and
# column of the window and of its first pane]]
function Parts($xl, $ws) {
  $notes = foreach ($c in $ws.Comments) { List @((Str $c.Parent.Address()), (Str ([string]$c.Author)), (Str ([string]$c.Text()))) }
  $threads = foreach ($t in $ws.CommentsThreaded) {
    $all = @($t) + @($t.Replies | ForEach-Object { $_ })
    List @((Str $t.Parent.Address()), (List ($all | ForEach-Object { List @((Str ([string]$_.Author.Name)), (Str ([string]$_.Text()))) })))
  }
  $links = foreach ($h in $ws.Hyperlinks) { List @((Str $h.Range.Address()), (Str ([string]$h.Address)), (Str ([string]$h.SubAddress)), (Str ([string]$h.ScreenTip)), (Str ([string]$h.TextToDisplay))) }
  [void]$ws.Activate()
  $w = $xl.ActiveWindow
  $p = $w.Panes.Item(1)
  $panes = List @((Cell $w.FreezePanes), (Cell ([double]$w.SplitRow)), (Cell ([double]$w.SplitColumn)), (Cell ([double]$w.ScrollRow)), (Cell ([double]$w.ScrollColumn)), (Cell ([double]$p.ScrollRow)), (Cell ([double]$p.ScrollColumn)))
  List @((Str $ws.Name), (List $notes), (List $threads), (List $links), $panes)
}

# what a parts case makes through the object model: a note, a hyperlink, frozen panes
function Make($xl, $wb, $s) {
  $ws = $wb.Worksheets.Item($s.sheet)
  switch ($s.kind) {
    'note' { [void]$ws.Range($s.cell).AddComment([string]$s.text) }
    'link' {
      $tip = if ($s.tip) { [string]$s.tip } else { [Type]::Missing }
      [void]$ws.Hyperlinks.Add($ws.Range($s.cell), [string]$s.url, [string]$s.location, $tip)
    }
    'freeze' {
      [void]$ws.Activate()
      $w = $xl.ActiveWindow
      $w.FreezePanes = $false
      $w.SplitRow = [int]$s.rows
      $w.SplitColumn = [int]$s.columns
      $w.FreezePanes = $true
    }
    default { throw "unknown make $($s.kind)" }
  }
}

# what a valid case makes through COM: a rule (Validation.Add with the range's first cell active, then its properties; flags
# "1", "0" or "" for Excel's own), a group of rows or columns, a collapsed group (ShowDetail), the summary settings
function VMake($wb, $s) {
  $ws = $wb.Worksheets.Item($s.sheet)
  [void]$ws.Activate()
  $m = [Type]::Missing
  switch ($s.kind) {
    'rule' {
      [void]$ws.Range($s.first).Select()
      $v = $ws.Range($s.range).Validation
      $f2 = if ($s.f2) { [string]$s.f2 } else { $m }
      if ($s.type -eq 0) { [void]$v.Add(0) } else { [void]$v.Add([int]$s.type, [int]$s.alert, [int]$s.operator, [string]$s.f1, $f2) }
      if ($s.ignore) { $v.IgnoreBlank = ($s.ignore -eq '1') }
      if ($s.dropdown) { $v.InCellDropdown = ($s.dropdown -eq '1') }
      if ($s.showInput) { $v.ShowInput = ($s.showInput -eq '1') }
      if ($s.showError) { $v.ShowError = ($s.showError -eq '1') }
      if ($s.inputTitle) { $v.InputTitle = [string]$s.inputTitle }
      if ($s.inputMessage) { $v.InputMessage = [string]$s.inputMessage }
      if ($s.errorTitle) { $v.ErrorTitle = [string]$s.errorTitle }
      if ($s.errorMessage) { $v.ErrorMessage = [string]$s.errorMessage }
    }
    'group' { if ($s.rows) { [void]$ws.Range($s.range).EntireRow.Group() } else { [void]$ws.Range($s.range).EntireColumn.Group() } }
    'collapse' { if ($s.rows) { $ws.Range($s.range).EntireRow.ShowDetail = $false } else { $ws.Range($s.range).EntireColumn.ShowDetail = $false } }
    'summary' {
      if ($s.below) { $ws.Outline.SummaryRow = if ($s.below -eq '1') { 1 } else { 0 } }
      if ($s.right) { $ws.Outline.SummaryColumn = if ($s.right -eq '1') { -4152 } else { -4131 } }
    }
    default { throw "unknown valid make $($s.kind)" }
  }
}

# an outline command (bench/windows/outline.ts) on lines `5:7` or `B:D` taken whole; Excel's refusal is the step's error
function OCommand($ws, $s) {
  $lines = if ($s.lines) { if ($s.lines -match '^\d') { , $ws.Range($s.lines).EntireRow } else { , $ws.Range($s.lines).EntireColumn } } else { $null }
  $m = [Type]::Missing
  switch ($s.kind) {
    'group' { [void]$lines.Group() }
    'ungroup' { [void]$lines.Ungroup() }
    'show' { $lines.ShowDetail = $true }
    'hide' { $lines.ShowDetail = $false }
    'hidden' { $lines.Hidden = $true }
    'shown' { $lines.Hidden = $false }
    'levels' { [void]$ws.Outline.ShowLevels($(if ($s.rows) { [int]$s.rows } else { $m }), $(if ($s.columns) { [int]$s.columns } else { $m })) }
    default { throw "unknown outline command $($s.kind)" }
  }
}

# the window's rows and columns that are grouped, hidden or summary lines, each [row or letters, level (0 for none),
# hidden, ShowDetail of a summary line or null]; and the tags of a copy of the book as Excel saves it: the first
# sheet's <row> tags with an outline level, hidden or collapsed, its <col>, <outlinePr> and <sheetFormatPr>
function ORead($wb, $s) {
  $ws = $wb.Worksheets.Item('S')
  $axes = foreach ($axis in 'rows', 'columns') {
    $span = $s.$axis
    $read = foreach ($k in [int]$span[0]..[int]$span[1]) {
      $e = if ($axis -eq 'rows') { $ws.Rows.Item($k) } else { $ws.Columns.Item($k) }
      $level = [int]$e.OutlineLevel - 1
      $summary = [bool]$e.Summary
      if ($level -gt 0 -or $e.Hidden -or $summary) {
        $key = if ($axis -eq 'rows') { Cell ([double]$k) } else { Str (($e.Address($false, $false) -split ':')[0]) }
        $detail = if ($summary) { try { Cell ([bool]$e.ShowDetail) } catch { 'null' } } else { 'null' }
        List @($key, (Cell ([double]$level)), (Cell ([bool]$e.Hidden)), $detail)
      }
    }
    , (List $read)
  }
  $copy = Join-Path $out 'copy.xlsx'
  if (Test-Path $copy) { Remove-Item -Force $copy }
  $wb.SaveCopyAs($copy)
  $zip = [IO.Compression.ZipFile]::OpenRead($copy)
  try {
    $r = New-Object IO.StreamReader($zip.GetEntry('xl/worksheets/sheet1.xml').Open(), [Text.Encoding]::UTF8)
    $xml = $r.ReadToEnd()
    $r.Close()
  } finally { $zip.Dispose() }
  $tags = foreach ($t in [regex]::Matches($xml, '<(row|col|outlinePr|sheetFormatPr)\b[^>]*>')) {
    if ($t.Groups[1].Value -ne 'row' -or $t.Value -match '\b(hidden|outlineLevel|collapsed)=') { Str $t.Value }
  }
  List @($axes[0], $axes[1], (List $tags))
}

# a property of a rule, or "" where COM has none for its type
function Prop($v, $name) { try { $v.$name } catch { '' } }

# a sheet's rules over the window, each cell selected (formulas are relative to the active cell), with Validation.Value;
# the window's grouped, hidden and collapsed lines; the summary settings: [name, [[address, type, operator, formula1,
# formula2, ignore blank, dropdown, show input, show error, alert, input title, input message, error title, error message,
# value]], [[row, level, hidden, collapsed]], [[column, level, hidden, collapsed]], [summary row, summary column]]
function VRead($ws, $win) {
  [void]$ws.Activate()
  $rules = @()
  $found = try { $ws.Range($win).SpecialCells(-4174) } catch { $null }
  if ($found) {
    foreach ($c in $found.Cells) {
      [void]$c.Select()
      $v = $c.Validation
      $rules += List @((Str $c.Address()), (Cell ([double]$v.Type)), (Cell ([double](Prop $v 'Operator'))), (Str ([string](Prop $v 'Formula1'))), (Str ([string](Prop $v 'Formula2'))), (Cell $v.IgnoreBlank), (Cell $v.InCellDropdown), (Cell $v.ShowInput), (Cell $v.ShowError), (Cell ([double]$v.AlertStyle)), (Str ([string]$v.InputTitle)), (Str ([string]$v.InputMessage)), (Str ([string]$v.ErrorTitle)), (Str ([string]$v.ErrorMessage)), (Cell $v.Value))
    }
  }
  $w = $ws.Range($win)
  $lines = foreach ($axis in 'rows', 'columns') {
    $out = foreach ($k in 1..$(if ($axis -eq 'rows') { $w.Rows.Count } else { $w.Columns.Count })) {
      $e = if ($axis -eq 'rows') { $ws.Rows.Item($w.Row + $k - 1) } else { $ws.Columns.Item($w.Column + $k - 1) }
      $level = [int]$e.OutlineLevel
      $collapsed = $false
      if ($e.Summary) { $collapsed = -not $e.ShowDetail }
      if ($level -gt 1 -or $e.Hidden -or $collapsed) {
        $key = if ($axis -eq 'rows') { Cell ([double]$e.Row) } else { Str (($e.Address($false, $false) -split ':')[0]) }
        List @($key, (Cell ([double]($level - 1))), (Cell $e.Hidden), (Cell $collapsed))
      }
    }
    , (List $out)
  }
  List @((Str $ws.Name), (List $rules), $lines[0], $lines[1], (List @((Cell ([double]$ws.Outline.SummaryRow)), (Cell ([double]$ws.Outline.SummaryColumn)))))
}

function Step($xl, $wb, $s) {
  switch ($s.do) {
    'full' { $xl.CalculateFull(); 'null' }
    'calc' { $xl.Calculate(); 'null' }
    'rebuild' { $xl.CalculateFullRebuild(); 'null' }
    'values' { Grid (Rg $wb $s.sheet $s.range).Value2 }
    'formulas' { Grid (Rg $wb $s.sheet $s.range).Formula }
    'formulas2' { Grid (Rg $wb $s.sheet $s.range).Formula2 }
    # the Sort command over a range by its first column, without a header (Range.Sort: Key1, Order1 1 ascending or 2
    # descending, Header 2 xlNo)
    'sort' {
      $rg = Rg $wb $s.sheet $s.range
      $order = if ($s.descending) { 2 } else { 1 }
      [void]$rg.Sort($rg.Cells.Item(1, 1), $order, [Type]::Missing, [Type]::Missing, 1, [Type]::Missing, 1, 2)
      'null'
    }
    # every cell of the range holds a formula: a book Excel repaired on opening lost its formulas
    'hasFormula' { Cell (Rg $wb $s.sheet $s.range).HasFormula }
    'shown' {
      $ws = $wb.Worksheets.Item($s.sheet)
      List ($s.cells | ForEach-Object { Shown $ws.Range($_) $s.codes })
    }
    'settings' { List @((Cell $xl.Iteration), (Cell ([double]$xl.MaxIterations)), (Cell $xl.MaxChange)) }
    # cells typed into: what each shows, its format codes and its prefix character (with `edit` its Formula and
    # FormulaLocal too, the text the formula bar opens it with); and Excel's automatic percent entry
    'entry' {
      $ws = $wb.Worksheets.Item($s.sheet)
      List ($s.cells | ForEach-Object {
        $c = $ws.Range($_)
        $read = @((Str ([string]$c.Text)), (Str $c.NumberFormat), (Str $c.NumberFormatLocal), (Str ([string]$c.PrefixCharacter)))
        if ($s.edit) { $read += @((Str ([string]$c.Formula)), (Str ([string]$c.FormulaLocal))) }
        List $read
      })
    }
    'autoPercent' { Cell $xl.AutoPercentEntry }
    # Find over a range or the whole sheet (bench/windows/find.ts): from after the range's last cell (of its last area),
    # so its first cell comes first, or after the cell the case names, the cells Find and FindNext give until they come back to one found before, at most
    # max, each [address, Formula, FormulaLocal, Text]; whether more were left, FindNext's error (over a range of one
    # cell), and the cell FindNext came back to when it is not the first (by columns on a filtered sheet)
    'find' {
      try {
        $ws = $wb.Worksheets.Item($s.sheet)
        # the comma keeps PowerShell from unrolling the range into its cells, the whole sheet's too; the areas of a range
        # are parted by the locale's list separator (`;` in ru-RU)
        $rg = if ($s.range) { , $ws.Range(([string]$s.range).Replace(',', $xl.International(5))) } else { , $ws.Cells }
        $a = $rg.Areas.Item($rg.Areas.Count)
        $last = if ($s.after) { $ws.Range([string]$s.after) } else { $a.Item($a.Rows.Count, $a.Columns.Count) }
        $lookIn = if ($s.lookIn -eq 'values') { -4163 } else { -4123 }
        $lookAt = if ($s.whole) { 1 } else { 2 }
        $order = if ($s.byColumns) { 2 } else { 1 }
        $f = $rg.Find([string]$s.what, $last, $lookIn, $lookAt, $order, 1, [bool]$s.matchCase, $false, $false)
        $first = if ($null -ne $f) { $f.Address($false, $false) } else { '' }
        $seen = @{}
        $found = @()
        $next = ''
        $back = ''
        while ($null -ne $f -and $found.Count -lt [int]$s.max) {
          $seen[$f.Address($false, $false)] = $true
          $found += List @((Str $f.Address($false, $false)), (Str ([string]$f.Formula)), (Str ([string]$f.FormulaLocal)), (Str ([string]$f.Text)))
          $f = try { $rg.FindNext($f) } catch { $next = "$($_.Exception.HResult) $($_.Exception.Message)"; $null }
          if ($null -ne $f -and $seen.ContainsKey($f.Address($false, $false))) {
            if ($f.Address($false, $false) -ne $first) { $back = $f.Address($false, $false) }
            $f = $null
          }
        }
        List @((List $found), (Cell ($null -ne $f)), (Str $next), (Str $back))
      } catch { Err $_ }
    }
    # Replace over a range or the whole sheet (bench/windows/replace.ts), as the dialog's Replace All with Look in
    # Formulas, a formula put back as a dynamic one (FormulaVersion xlReplaceFormula2, 1): what Range.Replace answers,
    # or its error; the range kept whole as Find's
    'replace' {
      try {
        $ws = $wb.Worksheets.Item($s.sheet)
        $rg = if ($s.range) { , $ws.Range(([string]$s.range).Replace(',', $xl.International(5))) } else { , $ws.Cells }
        $lookAt = if ($s.whole) { 1 } else { 2 }
        Cell ([bool]$rg.Replace([string]$s.what, [string]$s.by, $lookAt, 1, [bool]$s.matchCase, $false, $false, $false, 1))
      } catch { Err $_ }
    }
    # the cell Range.End comes to from a cell in a direction (XlDirection), its address (bench/windows/end.ts)
    'end' { try { Str (Rg $wb $s.sheet $s.cell).End([int]$s.direction).Address($false, $false) } catch { Err $_ } }
    # text put on the clipboard and pasted at a cell of the active sheet (bench/windows/ui.ts); the clipboard's or
    # Excel's refusal is the step's error
    'copyText' {
      $before = Get-Clipboard -Raw
      try {
        [void](Rg $wb $s.sheet $s.range).Copy()
        $copied = [string](Get-Clipboard -Raw)
        if ($s.at) {
          Set-Clipboard -Value $copied
          $ws = $wb.Worksheets.Item($s.sheet)
          [void]$ws.Activate()
          [void]$ws.Range($s.at).Select()
          [void]$ws.Paste()
          List @((Str $copied), (Grid (Rg $wb $s.sheet $s.read).Value2))
        } else { Str $copied }
      } finally {
        Set-Clipboard -Value ([string]$before)
      }
    }
    'paste' {
      try {
        Set-Clipboard -Value ([string]$s.text)
        $ws = $wb.Worksheets.Item($s.sheet)
        [void]$ws.Activate()
        [void]$ws.Range($s.at).Select()
        [void]$ws.Paste()
        'null'
      } catch { Err $_ }
    }
    # the cells of a ui case: what each shows, its format codes and prefix character, its number (Value2), its formulas
    # (Formula, FormulaLocal, Formula2, Formula2Local), its old-kind array (HasArray, FormulaArray) and its spill (HasSpill,
    # the address of SpillParent)
    'cells' {
      $ws = $wb.Worksheets.Item($s.sheet)
      List ($s.cells | ForEach-Object {
        $c = $ws.Range($_)
        $v = $c.Value2
        $num = if ($v -is [double]) { Cell $v } else { '""' }
        $fa = if ($c.HasArray) { Str ([string]$c.FormulaArray) } else { '""' }
        $sp = if ($c.HasSpill) { Str $c.SpillParent.Address($false, $false) } else { '""' }
        List @((Str ([string]$c.Text)), (Str $c.NumberFormat), (Str $c.NumberFormatLocal), (Str ([string]$c.PrefixCharacter)), $num, (Str ([string]$c.Formula)), (Str ([string]$c.FormulaLocal)), (Str ([string]$c.Formula2)), (Str ([string]$c.Formula2Local)), (Cell $c.HasArray), $fa, (Cell $c.HasSpill), $sp)
      })
    }
    # AutoFit of a range's whole columns; each column's ColumnWidth, Width (points) and UseStandardWidth (bench/windows/fit.ts)
    'autofit' { [void](Rg $wb $s.sheet $s.range).EntireColumn.AutoFit(); 'null' }
    'setColumnWidth' { (Rg $wb $s.sheet $s.column).ColumnWidth = [double]$s.width; 'null' }
    'standardWidth' { Cell ([double]$wb.Worksheets.Item($s.sheet).StandardWidth) }
    'widths' {
      $ws = $wb.Worksheets.Item($s.sheet)
      List ($s.columns | ForEach-Object { $c = $ws.Range("$($_)1"); List @((Cell ([double]$c.ColumnWidth)), (Cell ([double]$c.Width)), (Cell $c.UseStandardWidth)) })
    }
    # AutoFit of a range's whole rows; each row's RowHeight and UseStandardHeight, and cells' Formula and WrapText (bench/windows/rows.ts)
    'rowfit' { [void](Rg $wb $s.sheet $s.range).EntireRow.AutoFit(); 'null' }
    'heights' {
      $ws = $wb.Worksheets.Item($s.sheet)
      List ($s.rows | ForEach-Object { $r = $ws.Range("A$($_)"); List @((Cell ([double]$r.RowHeight)), (Cell $r.UseStandardHeight)) })
    }
    'wraps' {
      $ws = $wb.Worksheets.Item($s.sheet)
      List ($s.cells | ForEach-Object { $c = $ws.Range($_); List @((Str ([string]$c.Formula)), (Cell ([bool]$c.WrapText))) })
    }
    'set' {
      $rg = Rg $wb $s.sheet $s.cell
      try {
        switch ($s.kind) {
          'n' { $rg.Value2 = [double]::Parse($s.value, $inv) }
          's' { $rg.Value2 = [string]$s.value }
          'f' { $rg.Formula = [string]$s.value }
          'f2' { $rg.Formula2 = [string]$s.value }
          # an array formula of the old kind over the range (Ctrl+Shift+Enter)
          'fa' { $rg.FormulaArray = [string]$s.value }
          # typed text, as Excel 365 takes it in the user's language
          'l2' { $rg.Formula2Local = [string]$s.value }
          'x' { [void]$rg.ClearContents() }
        }
        'null'
      } catch { Err $_ }
    }
    # a number format, bold, italic or a fill (BGR) set on a range: a column, a row or cells (bench/windows/lines.ts)
    'format' {
      try {
        $rg = Rg $wb $s.sheet $s.range
        if ($s.numberFormat) { $rg.NumberFormat = [string]$s.numberFormat }
        if ($s.bold) { $rg.Font.Bold = $true }
        if ($s.italic) { $rg.Font.Italic = $true }
        if ($null -ne $s.fill) { $rg.Interior.Color = [double]$s.fill }
        'null'
      } catch { Err $_ }
    }
    # a range cleared, as ClearContents, ClearFormats or Clear (bench/windows/lines.ts)
    'clear' {
      try {
        $rg = Rg $wb $s.sheet $s.range
        switch ($s.what) {
          'contents' { [void]$rg.ClearContents() }
          'formats' { [void]$rg.ClearFormats() }
          'all' { [void]$rg.Clear() }
        }
        'null'
      } catch { Err $_ }
    }
    'hide' { (Rg $wb $s.sheet $s.range).EntireRow.Hidden = [bool]$s.hidden; 'null' }
    'hideColumns' { (Rg $wb $s.sheet $s.range).EntireColumn.Hidden = [bool]$s.hidden; 'null' }
    'filter' { [void](Rg $wb $s.sheet $s.range).AutoFilter([int]$s.field, [string]$s.criterion); 'null' }
    'struct' { try { Struct $wb $s; 'null' } catch { Err $_ } }
    'sheets' {
      List ($wb.Worksheets | ForEach-Object {
          $lk = if ($s.look) { Look $_ $s.look } else { '""' }
          List @((Str $_.Name), (Str $_.UsedRange.Address()), (Grid $_.Range($s.range).Formula), $lk)
        })
    }
    'formulaAt' {
      List ($s.cells | ForEach-Object {
          $f = try { $wb.Worksheets.Item($_.sheet).Range($_.cell).Formula } catch { '' }
          Str ([string]$f)
        })
    }
    # with visible, the names Excel shows only (it keeps hidden ones such as _xlfn. names)
    'names' { List ($wb.Names | Where-Object { -not $s.visible -or $_.Visible } | ForEach-Object { List @((Str $_.Name), (Str $_.RefersTo)) }) }
    'table' { try { Table $wb $s.line; 'null' } catch { Err $_ } }
    'formulasLocal' { Grid (Rg $wb $s.sheet $s.range).FormulaLocal }
    # every table of the sheets but those named in skip: [sheet, name, range, headers shown, totals shown, [columns]]
    'tables' {
      List ($wb.Worksheets | Where-Object { $s.skip -notcontains $_.Name } | ForEach-Object {
          $ws = $_
          $ws.ListObjects | ForEach-Object { List @((Str $ws.Name), (Str $_.Name), (Str $_.Range.Address()), (Cell $_.ShowHeaders), (Cell $_.ShowTotals), (List ($_.ListColumns | ForEach-Object { Str $_.Name }))) }
        })
    }
    'make' { try { Make $xl $wb $s; 'null' } catch { Err $_ } }
    'cond' { Cond $wb $s }
    'vmake' { try { VMake $wb $s; 'null' } catch { Err $_ } }
    'ocmd' { try { OCommand $wb.Worksheets.Item('S') $s; 'null' } catch { Err $_ } }
    'oread' { ORead $wb $s }
    'valid' {
      $all = List ($wb.Worksheets | ForEach-Object { VRead $_ $s.window })
      [void]$wb.Worksheets.Item(1).Activate()
      [void]$wb.Worksheets.Item(1).Range('A1').Select()
      $all
    }
    # a sheet's visibility through the object model: -1 visible, 0 hidden, 2 very hidden
    'visible' { try { $wb.Worksheets.Item($s.sheet).Visible = [int]$s.visible; 'null' } catch { Err $_ } }
    # the active sheet, and each sheet's name and visibility
    'book' { List @((Str $wb.ActiveSheet.Name), (List ($wb.Worksheets | ForEach-Object { List @((Str $_.Name), (Cell ([double]$_.Visible))) }))) }
    'parts' {
      $all = List ($wb.Worksheets | ForEach-Object { Parts $xl $_ })
      [void]$wb.Worksheets.Item(1).Activate()
      $all
    }
    # the book's folder and the linked books as Excel lists them (LinkSources of Excel links; none is Empty)
    'links' { List @((Str $wb.Path), (List @(@($wb.LinkSources(1)) | Where-Object { $null -ne $_ } | ForEach-Object { Str $_ }))) }
    # the book as Excel saves it (xlsx), beside its answer: out\<n>.saved.xlsx
    'saveAs' { $wb.SaveAs((Join-Path $out "$i.saved.xlsx"), 51); 'null' }
    default { throw "unknown step $($s.do)" }
  }
}

# the previous job's Excel may still be exiting, and a new one fails to start meanwhile (CO_E_SERVER_EXEC_FAILURE):
# wait for it, then start Excel, three times at most
for ($k = 0; $k -lt 30 -and @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq 0 }).Count; $k++) { Start-Sleep 1 }
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq 0 } | Stop-Process -Force
for ($k = 1; ; $k++) {
  try { $xl = New-Object -ComObject Excel.Application; break } catch { if ($k -ge 3) { throw }; Start-Sleep 10 }
}
$xl.DisplayAlerts = $false
try {
  for ($i = 0; $i -lt $job.books.Count; $i++) {
    $b = $job.books[$i]
    $t0 = Get-Date
    $res = try {
      if ($xl.Workbooks.Count -ne 0) { throw 'Excel has open books' }
      $wb = $xl.Workbooks.Open((Join-Path $dir $b.file), 0, $false)
      try {
        $answers = foreach ($s in $b.steps) { Step $xl $wb $s }
        '{"answers":' + (List $answers)
      } finally {
        $wb.Close($false)
      }
    } catch { '{"rejected":' + (Str "$($_.Exception.HResult) $($_.Exception.Message)") }
    $res += ',"seconds":' + ((Get-Date) - $t0).TotalSeconds.ToString('0.000', $inv) + '}'
    [IO.File]::WriteAllText((Join-Path $out "$i.json"), $res, [Text.Encoding]::ASCII)
  }
} finally {
  $left = $xl.Workbooks.Count
  $xl.Quit()
  [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  [IO.File]::WriteAllText((Join-Path $out 'done.json'), "{`"booksLeft`":$left}", [Text.Encoding]::ASCII)
}
