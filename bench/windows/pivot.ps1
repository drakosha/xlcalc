# Pivot tables for xl.ps1 (bench/windows/pivot.ts): each step one primitive of the object model, its numbers Excel's
# constants as bench/coverage/pivot/com.ts gives them, so the cases stay data. Dot-sourced by xl.ps1, whose Str, Cell,
# Grid, List and Err it uses.

# an argument from JSON: Missing for null, { d: serial } a date, a whole number an Int32 (COM takes no Int64 for an enum)
function PArg($x) {
  if ($null -eq $x) { return [Type]::Missing }
  if ($x -is [pscustomobject] -and $null -ne $x.d) { return [DateTime]::FromOADate([double]$x.d) }
  if ($x -is [long] -or $x -is [int]) { return [int]$x }
  if ($x -is [decimal]) { return [double]$x }
  if ($x -is [array]) { return [object[]]@($x | ForEach-Object { PArg $_ }) }
  $x
}

function PTable($wb, $s) { $wb.Worksheets.Item([string]$s.sheet).PivotTables([string]$s.pt) }

# what a step works on: the pivot table, one of its fields, a data field by its caption, the values field (Σ Values),
# a field's cube field, a range of a sheet, an item, the cache, a slicer cache by name or one of its items, a slicer; the comma keeps a collection whole
function POn($wb, $s) {
  switch ($s.on) {
    'table' { return , (PTable $wb $s) }
    'field' { return , (PTable $wb $s).PivotFields([string]$s.field) }
    'data' { return , (PTable $wb $s).DataFields([string]$s.field) }
    'values' { return , (PTable $wb $s).DataPivotField }
    'cubeField' { return , (PTable $wb $s).PivotFields([string]$s.field).CubeField }
    'range' { return , $wb.Worksheets.Item([string]$s.rangeSheet).Range([string]$s.range) }
    'item' { return , (PTable $wb $s).PivotFields([string]$s.field).PivotItems([string]$s.item) }
    'cache' { return , (PTable $wb $s).PivotCache() }
    'calcFields' { return , (PTable $wb $s).CalculatedFields() }
    'calcItems' { return , (PTable $wb $s).PivotFields([string]$s.field).CalculatedItems() }
    'slicerCache' { return , $wb.SlicerCaches.Item([string]$s.field) }
    'slicerItem' { return , $wb.SlicerCaches.Item([string]$s.field).SlicerItems([string]$s.item) }
    'timeline' { return , $wb.SlicerCaches.Item([string]$s.field).TimelineState }
    'timelineView' { return , $wb.SlicerCaches.Item([string]$s.field).Slicers.Item(1).TimelineViewState }
    default { throw "unknown pivot target $($s.on)" }
  }
}

# a step's arguments as an array of plain values: values that came through the pipeline are wrapped, and a COM method
# called through InvokeMember takes a wrapped string as a type mismatch
function PArgs($xs) {
  $a = New-Object 'object[]' @($xs).Count
  for ($k = 0; $k -lt $a.Count; $k++) { $v = PArg @($xs)[$k]; $a[$k] = if ($v -is [psobject]) { $v.PSObject.BaseObject } else { $v } }
  , $a
}

# the book's pivot tables, [sheet, name] each
function PAll($wb) {
  foreach ($ws in $wb.Worksheets) { foreach ($pt in $ws.PivotTables()) { , @([string]$ws.Name, [string]$pt.Name) } }
}

$PInvoke = [Reflection.BindingFlags]::InvokeMethod
$PPut = [Reflection.BindingFlags]::SetProperty

# a range of the pivot table by its property's name (DataBodyRange, PageRange…), its address; empty when there is none
function PRange($pt, $name) { try { $r = $pt.$name; if ($null -eq $r) { '' } else { [string]$r.Address($false, $false) } } catch { '' } }

# the pivot table's area: TableRange1, TableRange2, DataBodyRange (empty when there is none), each cell of TableRange2
# by rows (Value2, Text and NumberFormat), then PageRange, RowRange and ColumnRange (empty when there is none)
function PRead($wb, $s) {
  $pt = PTable $wb $s
  $r2 = $pt.TableRange2
  $body = PRange $pt 'DataBodyRange'
  $texts = for ($r = 1; $r -le $r2.Rows.Count; $r++) {
    $row = for ($c = 1; $c -le $r2.Columns.Count; $c++) { Str ([string]$r2.Cells.Item($r, $c).Text) }
    List $row
  }
  $formats = for ($r = 1; $r -le $r2.Rows.Count; $r++) {
    $row = for ($c = 1; $c -le $r2.Columns.Count; $c++) { Str ([string]$r2.Cells.Item($r, $c).NumberFormat) }
    List $row
  }
  List @((Str $pt.TableRange1.Address($false, $false)), (Str $r2.Address($false, $false)), (Str $body), (Grid $r2.Value2), (List $texts), (List $formats), (Str (PRange $pt 'PageRange')), (Str (PRange $pt 'RowRange')), (Str (PRange $pt 'ColumnRange')))
}

# a slicer cache's items as Excel shows them, [name, selected, has data]; a timeline's state, { timeline: [start, end
# (serials, null without a filter), level] }
function PSlicerRead($wb, $s) {
  $sc = $wb.SlicerCaches.Item([string]$s.field)
  if ($sc.SlicerCacheType -eq 2) {
    $t = $sc.TimelineState
    $range = if ($t.SingleRangeFilterState) { @((Cell ([double]$t.StartDate.ToOADate())), (Cell ([double]$t.EndDate.ToOADate()))) } else { @('null', 'null') }
    return '{"timeline":' + (List ($range + @(Cell ([double]$sc.Slicers.Item(1).TimelineViewState.Level)))) + '}'
  }
  List ($sc.SlicerItems | ForEach-Object { List @((Str ([string]$_.Name)), (Cell ([bool]$_.Selected)), (Cell ([bool]$_.HasData))) })
}

# the cells an external step names ("S!A1"), each one's Value2, in their order
function PCells($book, $s) {
  List ($s.cells.PSObject.Properties | ForEach-Object { $p = $_.Name -split '!'; Cell $book.Worksheets.Item($p[0]).Range($p[1]).Value2 })
}

# formulas over a pivot table of another book (source, base64), opened beside as Src.xlsx: typed and read with it open;
# the book saved as Main.xlsx; read with Src.xlsx closed, then computed fully and read; Main.xlsx opened alone without
# updating its links and read, computed fully and read; then opened after Src.xlsx, computed and read. Each read is
# [label, [values]]; an Excel error ends the reads with it.
function PExternal($xl, $wb, $s) {
  $sub = Join-Path $dir "x$i"
  New-Item -ItemType Directory -Force $sub | Out-Null
  $src = Join-Path $sub 'Src.xlsx'
  $main = Join-Path $sub 'Main.xlsx'
  [IO.File]::WriteAllBytes($src, [Convert]::FromBase64String($s.source))
  $reads = @()
  $other = $null
  $m = $null
  try {
    $other = $xl.Workbooks.Open($src, 0, $true)
    foreach ($p in $s.cells.PSObject.Properties) { $a = $p.Name -split '!'; $wb.Worksheets.Item($a[0]).Range($a[1]).Formula2 = [string]$p.Value }
    $xl.CalculateFull()
    $reads += List @((Str 'open'), (PCells $wb $s))
    $wb.SaveCopyAs($main)
    $other.Close($false); $other = $null
    $reads += List @((Str 'closed'), (PCells $wb $s))
    $xl.CalculateFull()
    $reads += List @((Str 'closed, computed'), (PCells $wb $s))
    $m = $xl.Workbooks.Open($main, 0, $true)
    $reads += List @((Str 'reopened'), (PCells $m $s))
    $xl.CalculateFull()
    $reads += List @((Str 'reopened, computed'), (PCells $m $s))
    $m.Close($false); $m = $null
    $other = $xl.Workbooks.Open($src, 0, $true)
    $m = $xl.Workbooks.Open($main, 0, $true)
    $xl.CalculateFull()
    $reads += List @((Str 'with the source'), (PCells $m $s))
    List $reads
  } catch {
    List ($reads + @(Err $_))
  } finally {
    if ($m) { $m.Close($false) }
    if ($other) { $other.Close($false) }
  }
}

function Pivot($xl, $wb, $s) {
  try {
    switch ($s.op) {
      # a cache over the source (a range with its sheet, a table's or a defined name) and a pivot table from it, or a
      # pivot table from the cache of another (cacheOf, on cacheSheet)
      'create' {
        $dest = $wb.Worksheets.Item([string]$s.sheet).Range([string]$s.at)
        $pc = if ($s.cacheOf) { $wb.Worksheets.Item([string]$s.cacheSheet).PivotTables([string]$s.cacheOf).PivotCache() } else { $wb.PivotCaches().Create(1, [string]$s.source, [int]$s.version) }
        [void]$pc.CreatePivotTable($dest, [string]$s.pt, $true, [int]$s.version)
        'null'
      }
      # a data field (AddDataField with its function), then its properties in order
      'data' {
        $pt = PTable $wb $s
        $df = $pt.AddDataField($pt.PivotFields([string]$s.field), (PArg $s.caption), [int]$s.fn)
        foreach ($p in $s.props) { $df.([string]$p[0]) = PArg $p[1] }
        Str ([string]$df.Caption)
      }
      # a data field with the field alone, the function Excel gives it: its caption and XlConsolidationFunction
      'dataDefault' {
        $pt = PTable $wb $s
        $df = $pt.AddDataField($pt.PivotFields([string]$s.field))
        List @((Str ([string]$df.Caption)), (Cell ([double]$df.Function)))
      }
      # properties set in order on the step's target
      'set' {
        $o = POn $wb $s
        foreach ($p in $s.props) { $o.([string]$p[0]) = PArg $p[1] }
        'null'
      }
      # a method of the target with its arguments
      'call' {
        $o = POn $wb $s
        $r = [System.__ComObject].InvokeMember([string]$s.method, $PInvoke, $null, $o, (PArgs $s.args))
        if ($null -eq $r -or $r -is [__ComObject]) { 'null' } else { Cell $r }
      }
      # a property with an index set (Subtotals(i) = v), each pair in order
      'put' {
        $o = POn $wb $s
        foreach ($p in $s.props) { [void][System.__ComObject].InvokeMember([string]$s.name, $PPut, $null, $o, (PArgs $p)) }
        'null'
      }
      # a label, value or date filter (PivotFilters.Add2), over a data field by its caption where given
      'filter' {
        $pt = PTable $wb $s
        $df = if ($s.dataField) { $pt.DataFields([string]$s.dataField) } else { [Type]::Missing }
        [void]$pt.PivotFields([string]$s.field).PivotFilters.Add2([int]$s.type, $df, (PArg $s.value1), (PArg $s.value2), [Type]::Missing, [Type]::Missing, [Type]::Missing, [Type]::Missing, (PArg $s.wholeDays))
        'null'
      }
      # Range.Group on the field's first item cell: start and end (true for Excel's own), the interval, the periods
      # (seconds, minutes, hours, days, months, quarters, years)
      'group' {
        $f = (PTable $wb $s).PivotFields([string]$s.field)
        $periods = if ($s.periods) { [object[]]@($s.periods | ForEach-Object { [bool]$_ }) } else { [Type]::Missing }
        [void]$f.DataRange.Cells.Item(1).Group((PArg $s.start), (PArg $s.end), (PArg $s.by), $periods)
        'null'
      }
      # a group of the items named: their label cells as one range (its areas parted by the locale's list separator, as
      # Application.Union takes no two ranges from PowerShell) and grouped
      'groupItems' {
        $pt = PTable $wb $s
        $f = $pt.PivotFields([string]$s.field)
        $addr = @($s.items | ForEach-Object { $f.PivotItems([string]$_).LabelRange.Address($false, $false) }) -join $xl.International(5)
        [void]$pt.TableRange1.Worksheet.Range($addr).Group()
        'null'
      }
      # only the items named shown (a field's PivotItems, Visible) or selected (a slicer's SlicerItems, Selected), compared
      # by name with case: the named ones first, as Excel refuses to hide every item
      'only' {
        $slicer = $s.on -eq 'slicerCache'
        $all = if ($slicer) { @($wb.SlicerCaches.Item([string]$s.field).SlicerItems) } else { @((PTable $wb $s).PivotFields([string]$s.field).PivotItems()) }
        $prop = if ($slicer) { 'Selected' } else { 'Visible' }
        $keep = @($s.items | ForEach-Object { [string]$_ })
        foreach ($it in $all) { if ($keep -ccontains [string]$it.Name -and -not $it.$prop) { $it.$prop = $true } }
        foreach ($it in $all) { if ($keep -cnotcontains [string]$it.Name -and $it.$prop) { $it.$prop = $false } }
        'null'
      }
      # Ungroup on the first item cell of the field on the rows or columns: Excel 365 names a date field grouped into
      # several `Months (Date)` and `Years (Date)`, so a cell of a field named so is taken too
      'ungroup' {
        $pt = PTable $wb $s
        $name = [string]$s.field
        $c = $null
        foreach ($axis in @($pt.RowRange, $pt.ColumnRange)) {
          if ($null -eq $axis) { continue }
          foreach ($x in $axis.Cells) {
            $pc = try { $x.PivotCell } catch { $null }
            if ($pc -and $pc.PivotCellType -eq 1 -and ($pc.PivotField.Name -eq $name -or $pc.PivotField.Name.EndsWith(" ($name)"))) { $c = $x; break }
          }
          if ($c) { break }
        }
        [void]$c.Ungroup()
        'null'
      }
      # a slicer cache (SlicerCaches.Add2, xlTimeline 2 for a timeline) on a field, a slicer of it on a sheet, the other
      # pivot tables it filters
      'slicer' {
        $pt = PTable $wb $s
        $sc = $wb.SlicerCaches.Add2($pt, [string]$s.field, [string]$s.name, [int]$s.kind)
        [void]$sc.Slicers.Add($wb.Worksheets.Item([string]$s.at), [Type]::Missing, [string]$s.name, [string]$s.field, 10, 400, 150, 200)
        foreach ($o in $s.also) { $sc.PivotTables.AddPivotTable($wb.Worksheets.Item([string]$o[0]).PivotTables([string]$o[1])) }
        Str ([string]$sc.Name)
      }
      # the table of the sheet D named resized to a range there
      'resize' { [void]$wb.Worksheets.Item('D').ListObjects.Item([string]$s.table).Resize($wb.Worksheets.Item('D').Range([string]$s.range)); 'null' }
      # a pivot table connected to the slicer cache named, or disconnected
      'link' {
        $sc = $wb.SlicerCaches.Item([string]$s.field)
        $pt = PTable $wb $s
        if ($s.off) { [void]$sc.PivotTables.RemovePivotTable($pt) } else { [void]$sc.PivotTables.AddPivotTable($pt) }
        'null'
      }
      # what RefreshTable answers (true when Excel refreshed)
      'drill' {
        $before = @($wb.Worksheets | ForEach-Object { [string]$_.Name })
        $wb.Worksheets.Item([string]$s.sheet).Range([string]$s.cell).ShowDetail = $true
        $added = @($wb.Worksheets | Where-Object { $before -notcontains [string]$_.Name })
        if ($added.Count -ne 1) { throw 'Show Details did not make one sheet' }
        Grid $added[0].UsedRange.Value2
      }
      # a range copied and pasted at a cell of P with it selected (Ctrl+C, Ctrl+V); the copy's marquee stays, as COM takes
      # no False for CutCopyMode from PowerShell
      'copyPaste' {
        $p = ([string]$s.from) -split '!'
        [void]$wb.Worksheets.Item($p[0]).Range($p[1]).Copy()
        $ws = $wb.Worksheets.Item('P')
        [void]$ws.Activate()
        [void]$ws.Range([string]$s.at).Select()
        try { [void]$ws.Paste() } finally { [void]$wb.Worksheets.Item(1).Activate() }
        'null'
      }
      'refresh' { Cell ([bool](PTable $wb $s).RefreshTable()) }
      # every pivot table of the book, [sheet, name], in the order of the sheets
      'list' { List (PAll $wb | ForEach-Object { List @((Str $_[0]), (Str $_[1])) }) }
      # each pivot table of the book read, [sheet, name, its area]
      'readAll' { List (PAll $wb | ForEach-Object { List @((Str $_[0]), (Str $_[1]), (PRead $wb @{ sheet = $_[0]; pt = $_[1] })) }) }
      # each pivot table of the book refreshed, [sheet, name, what RefreshTable answered or its error]
      'refreshAll' {
        List (PAll $wb | ForEach-Object {
            $r = try { Cell ([bool]$wb.Worksheets.Item($_[0]).PivotTables($_[1]).RefreshTable()) } catch { Err $_ }
            List @((Str $_[0]), (Str $_[1]), $r)
          })
      }
      'read' { PRead $wb $s }
      'external' { PExternal $xl $wb $s }
      'slicerRead' { PSlicerRead $wb $s }
      # a copy of the book as it is now, beside its answer: out\<n>.copy<k>.xlsx
      'copy' { $wb.SaveCopyAs((Join-Path $out "$i.copy$($s.k).xlsx")); 'null' }
      default { throw "unknown pivot op $($s.op)" }
    }
  } catch { Err $_ }
}
