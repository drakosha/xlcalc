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

# the look of a sheet: cells [address, bold, fill BGR, format, merged area], columns [letter, standard, width, hidden], rows likewise
function Look($ws, $s) {
  $cells = foreach ($a in $s.cells) {
    $c = $ws.Range($a)
    $m = if ($c.MergeCells) { $c.MergeArea.Address() } else { '' }
    List @((Str $a), (Cell $c.Font.Bold), (Cell $c.Interior.Color), (Str $c.NumberFormat), (Str $m))
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

function Step($xl, $wb, $s) {
  switch ($s.do) {
    'full' { $xl.CalculateFull(); 'null' }
    'calc' { $xl.Calculate(); 'null' }
    'values' { Grid (Rg $wb $s.sheet $s.range).Value2 }
    'formulas' { Grid (Rg $wb $s.sheet $s.range).Formula }
    'formulas2' { Grid (Rg $wb $s.sheet $s.range).Formula2 }
    # every cell of the range holds a formula: a book Excel repaired on opening lost its formulas
    'hasFormula' { Cell (Rg $wb $s.sheet $s.range).HasFormula }
    'shown' {
      $ws = $wb.Worksheets.Item($s.sheet)
      List ($s.cells | ForEach-Object { Shown $ws.Range($_) $s.codes })
    }
    'settings' { List @((Cell $xl.Iteration), (Cell ([double]$xl.MaxIterations)), (Cell $xl.MaxChange)) }
    'set' {
      $rg = Rg $wb $s.sheet $s.cell
      try {
        switch ($s.kind) {
          'n' { $rg.Value2 = [double]::Parse($s.value, $inv) }
          's' { $rg.Value2 = [string]$s.value }
          'f' { $rg.Formula = [string]$s.value }
          'f2' { $rg.Formula2 = [string]$s.value }
          'x' { [void]$rg.ClearContents() }
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
    'names' { List ($wb.Names | ForEach-Object { List @((Str $_.Name), (Str $_.RefersTo)) }) }
    default { throw "unknown step $($s.do)" }
  }
}

$xl = New-Object -ComObject Excel.Application
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
