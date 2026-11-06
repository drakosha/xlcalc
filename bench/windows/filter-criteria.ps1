# A filter's criteria and sort as Excel for Windows makes and saves them, one book a case: every kind on a sheet's AutoFilter
# and on a table's, a sheet's own sort, and what Excel does to them when lines go, cells move, a sheet is copied.
param([string]$dir)
$ErrorActionPreference = 'Continue'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
$xl.Visible = $false
$m = [Type]::Missing
$out = @()
$RED = 255
$BLUE = 16711680
$LONG = 'L' + ('x' * 299)

# Region, Product, Amount, When, Score, Note over A1:F21: Amount's cells filled and coloured by turns, Score under an icon set
function Fill($ws) {
  $regions = 'North', 'South', 'East', 'West'
  $products = 'Apple', 'apple', 'Banana', 'Cherry', 'Date'
  $head = 'Region', 'Product', 'Amount', 'When', 'Score', 'Note'
  for ($c = 0; $c -lt 6; $c++) { $ws.Cells.Item(1, $c + 1).Value2 = $head[$c] }
  for ($k = 0; $k -lt 20; $k++) {
    $r = $k + 2
    if ($k % 7 -ne 3) { $ws.Cells.Item($r, 1).Value2 = $regions[$k % 4] }
    $ws.Cells.Item($r, 2).Value2 = $products[$k % 5]
    $ws.Cells.Item($r, 3).Formula = [string](5 * ($k + 1))
    if ($k % 4 -eq 1) { $ws.Cells.Item($r, 3).Interior.Color = $RED }
    if ($k % 4 -eq 2) { $ws.Cells.Item($r, 3).Font.Color = $BLUE }
    # two years, months and days apart, hours and minutes apart, seconds apart
    $ws.Cells.Item($r, 4).NumberFormat = 'yyyy-mm-dd hh:mm:ss'
    $ws.Cells.Item($r, 4).Formula = ([double](45000 + 40 * $k + ($k % 3) * 0.25 + ($k % 2) * 0.01 + ($k % 5) / 86400)).ToString('R', [Globalization.CultureInfo]::InvariantCulture)
    $ws.Cells.Item($r, 5).Formula = [string]($k + 1)
    if ($k % 9 -eq 4) { $ws.Cells.Item($r, 6).Value2 = $LONG } elseif ($k % 2) { $ws.Cells.Item($r, 6).Value2 = "note $k" }
  }
  [void]$ws.Range('E2:E21').FormatConditions.AddIconSetCondition()
  $ws.Columns.Item(4).ColumnWidth = 20
}

function Hidden($ws) {
  $list = @()
  $last = [Math]::Max(25, $ws.UsedRange.Row + $ws.UsedRange.Rows.Count)
  for ($r = 1; $r -le $last; $r++) { if ($ws.Rows.Item($r).Hidden) { $list += $r } }
  return , $list
}

# Case: $kind sheet or table. $act gets the book, the sheet, the range the filter goes over and the sort's owner.
function Case($id, $kind, [scriptblock]$act) {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'Data'
  Fill $ws
  $err = ''
  $lo = $null
  try {
    if ($kind -eq 'table') {
      $lo = $ws.ListObjects.Add(1, $ws.Range('A1:F21'), $m, 1)
      $lo.Name = 'Sales'
      & $act $wb $ws $lo.Range $lo
    } else {
      $r = $ws.Range('A1:F21')
      [void]$r.AutoFilter()
      & $act $wb $ws $r $null
    }
  } catch { $err = $_.Exception.Message }
  try { $wb.SaveAs("$dir\$id.xlsx", 51) } catch { $err += ' SAVE ' + $_.Exception.Message }
  $sh = $wb.Worksheets.Item(1)
  $range = ''
  try { if ($sh.AutoFilter) { $range = [string]$sh.AutoFilter.Range.Address($false, $false) } } catch {}
  $script:out += @{ id = $id; kind = $kind; refused = $err; filterMode = [bool]$sh.FilterMode; range = $range; hidden = (Hidden $sh) }
  $wb.Close($false)
}

# Sort: the sheet's AutoFilter's or the table's, by one field
function SortOf($ws, $lo) { if ($lo) { return $lo.Sort } else { return $ws.AutoFilter.Sort } }
function Apply($s) { $s.Header = 1; $s.Apply() }

# every kind of criterion and sort, on a sheet's filter and on a table's
foreach ($kind in 'sheet', 'table') {
  $p = "$kind-"
  Case "${p}plain" $kind { param($wb, $ws, $r, $lo) }
  Case "${p}values" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, [string[]]@('Banana', 'Cherry'), 7) }
  Case "${p}one-value" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana') }
  Case "${p}numbers" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, [string[]]@('10', '20', '35'), 7) }
  Case "${p}values-blank" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(1, [string[]]@('North', '='), 7) }
  Case "${p}blanks" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(1, '=') }
  Case "${p}non-blanks" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(1, '<>') }
  Case "${p}year" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, $m, 7, [object[]]@(0, '6/1/2024')) }
  Case "${p}month" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, $m, 7, [object[]]@(1, '4/9/2023')) }
  Case "${p}day" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, $m, 7, [object[]]@(2, '3/15/2024')) }
  Case "${p}hour" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, $m, 7, [object[]]@(3, '3/15/2024 6:00:00')) }
  Case "${p}minute" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, $m, 7, [object[]]@(4, '3/15/2024 6:14:00')) }
  Case "${p}second" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, $m, 7, [object[]]@(5, '3/15/2024 6:14:24')) }
  Case "${p}dates-two" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, $m, 7, [object[]]@(0, '6/1/2023', 1, '4/9/2024')) }
  Case "${p}cond-and" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, '>20', 1, '<=70') }
  Case "${p}cond-or" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, '<10', 2, '>=90') }
  Case "${p}cond-text" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, '=*an*') }
  Case "${p}cond-not" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, '<>Apple') }
  Case "${p}cond-one" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, '>=50') }
  Case "${p}top" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, '5', 3) }
  Case "${p}top-percent" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, '25', 5) }
  Case "${p}bottom" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, '3', 4) }
  Case "${p}bottom-percent" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, '10', 6) }
  Case "${p}above-average" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, 33, 11) }
  Case "${p}below-average" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, 34, 11) }
  Case "${p}today" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, 1, 11) }
  Case "${p}this-month" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, 7, 11) }
  Case "${p}last-year" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, 14, 11) }
  Case "${p}year-to-date" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, 16, 11) }
  Case "${p}q1" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, 17, 11) }
  Case "${p}m3" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(4, 23, 11) }
  Case "${p}fill" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, $RED, 8) }
  Case "${p}font" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, $BLUE, 9) }
  Case "${p}no-fill" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, $m, 12) }
  Case "${p}auto-font" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, $m, 13) }
  Case "${p}icon" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(5, $wb.IconSets.Item(1).Item(1), 10) }
  Case "${p}icon-third" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(5, $wb.IconSets.Item(1).Item(3), 10) }
  Case "${p}no-icon" $kind { param($wb, $ws, $r, $lo) $ws.Range('E2:E5').FormatConditions.Delete(); [void]$r.AutoFilter(5, $m, 14) }
  Case "${p}hidden-button" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, $m, 1, $m, $false) }
  Case "${p}hidden-button-filter" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana', 1, $m, $false) }
  Case "${p}long-value" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(6, [string[]]@($LONG), 7) }
  Case "${p}long-values" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(6, [string[]]@($LONG, 'note 1'), 7) }
  Case "${p}long-cond" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(6, '=' + $LONG) }
  Case "${p}long-not" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(6, '<>' + $LONG) }
  Case "${p}long-wild" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(6, '=*' + $LONG) }
  Case "${p}long-two" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(6, '=' + $LONG, 2, '=note 1') }
  Case "${p}two-columns" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(1, [string[]]@('North', 'South'), 7); [void]$r.AutoFilter(3, '>40') }
  Case "${p}sort-value" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 1); Apply $s }
  Case "${p}sort-desc" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s }
  Case "${p}sort-two" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('A2:A21'), 0, 1); [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s }
  Case "${p}sort-fill" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; $f = $s.SortFields.Add($ws.Range('C2:C21'), 1, 1); $f.SortOnValue.Color = $RED; Apply $s }
  Case "${p}sort-font" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; $f = $s.SortFields.Add($ws.Range('C2:C21'), 2, 1); $f.SortOnValue.Color = $BLUE; Apply $s }
  Case "${p}sort-icon" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; $f = $s.SortFields.Add($ws.Range('E2:E21'), 3, 1); $f.SetIcon($wb.IconSets.Item(1).Item(1)); Apply $s }
  Case "${p}sort-list" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('A2:A21'), 0, 1, 'West,East,South,North'); Apply $s }
  Case "${p}sort-case" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; $s.MatchCase = $true; [void]$s.SortFields.Add($ws.Range('B2:B21'), 0, 1); Apply $s }
  Case "${p}sort-stroke" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; $s.SortMethod = 2; [void]$s.SortFields.Add($ws.Range('B2:B21'), 0, 1); Apply $s }
  Case "${p}sort-pinyin" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; $s.SortMethod = 1; [void]$s.SortFields.Add($ws.Range('B2:B21'), 0, 1); Apply $s }
  Case "${p}sort-filter" $kind { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, [string[]]@('Banana', 'Cherry'), 7); $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s }
  # what lines and moves do to them
  $filtered = { param($r) [void]$r.AutoFilter(2, [string[]]@('Banana', 'Cherry'), 7); [void]$r.AutoFilter(3, '>20') }
  Case "${p}column-before" $kind { param($wb, $ws, $r, $lo) & $filtered $r; [void]$ws.Columns.Item(2).Insert() }
  Case "${p}column-between" $kind { param($wb, $ws, $r, $lo) & $filtered $r; [void]$ws.Columns.Item(3).Insert() }
  Case "${p}column-deleted" $kind { param($wb, $ws, $r, $lo) & $filtered $r; [void]$ws.Columns.Item(2).Delete() }
  Case "${p}column-other-deleted" $kind { param($wb, $ws, $r, $lo) & $filtered $r; [void]$ws.Columns.Item(1).Delete() }
  Case "${p}row-inside" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; [void]$ws.Rows.Item(5).Insert() }
  Case "${p}row-deleted" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; [void]$ws.Rows.Item(5).Delete() }
  Case "${p}row-above" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; [void]$ws.Rows.Item(1).Insert() }
  Case "${p}sort-column-deleted" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; [void]$ws.Columns.Item(3).Delete() }
  Case "${p}sort-one-of-two-deleted" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('A2:A21'), 0, 1); [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; [void]$ws.Columns.Item(3).Delete() }
  Case "${p}sort-column-before" $kind { param($wb, $ws, $r, $lo) $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; [void]$ws.Columns.Item(2).Insert() }
  Case "${p}moved" $kind { param($wb, $ws, $r, $lo) & $filtered $r; $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; [void]$ws.Range('A1:F21').Cut($ws.Range('H3')) }
  Case "${p}sheet-copied" $kind { param($wb, $ws, $r, $lo) & $filtered $r; $s = SortOf $ws $lo; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; $ws.Copy($m, $ws) }
  Case "${p}sheet-renamed" $kind { param($wb, $ws, $r, $lo) & $filtered $r; $ws.Name = 'Goods' }
}

# a sheet's own: its filter's further forms
Case 'sheet-merged-header' 'sheet' { param($wb, $ws, $r, $lo) $ws.AutoFilterMode = $false; $ws.Range('A1:B1').Merge(); $x = $ws.Range('A1:F21'); [void]$x.AutoFilter() }
Case 'sheet-merged-filter' 'sheet' { param($wb, $ws, $r, $lo) $ws.AutoFilterMode = $false; $ws.Range('A1:B1').Merge(); $x = $ws.Range('A1:F21'); [void]$x.AutoFilter(1, 'North') }
Case 'sheet-header-deleted' 'sheet' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); [void]$ws.Rows.Item(1).Delete() }
Case 'sheet-moved-other' 'sheet' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); $o = $wb.Worksheets.Add(); $o.Name = 'Other'; [void]$ws.Range('A1:F21').Cut($o.Range('A1')) }
Case 'sheet-moved-part' 'sheet' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); [void]$ws.Range('A1:B21').Cut($ws.Range('H3')) }
Case 'sheet-table-over' 'sheet' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, [string[]]@('Banana', 'Cherry'), 7); $s = $ws.AutoFilter.Sort; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; [void]$ws.ListObjects.Add(1, $ws.Range('A1:F21'), $m, 1) }
Case 'sheet-table-over-shown' 'sheet' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, $m, 1, $m, $false); [void]$ws.ListObjects.Add(1, $ws.Range('A1:F21'), $m, 1) }
Case 'sheet-table-over-more' 'sheet' { param($wb, $ws, $r, $lo) $ws.AutoFilterMode = $false; $x = $ws.Range('B1:E21'); [void]$x.AutoFilter(1, [string[]]@('Banana', 'Cherry'), 7); $s = $ws.AutoFilter.Sort; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; [void]$ws.ListObjects.Add(1, $ws.Range('A1:F21'), $m, 1) }
Case 'sheet-moved-button' 'sheet' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, $m, 1, $m, $false); [void]$ws.Range('A1:F21').Cut($ws.Range('H3')) }
Case 'sheet-moved-rows' 'sheet' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); [void]$ws.Rows.Item('1:21').Cut($ws.Rows.Item('30:50')) }
Case 'sheet-icon-stars' 'sheet' { param($wb, $ws, $r, $lo) $c = $ws.Range('E2:E21').FormatConditions.Item(1); $c.IconSet = $wb.IconSets.Item(18); [void]$r.AutoFilter(5, $wb.IconSets.Item(18).Item(1), 10) }
Case 'sheet-sort-stars' 'sheet' { param($wb, $ws, $r, $lo) $c = $ws.Range('E2:E21').FormatConditions.Item(1); $c.IconSet = $wb.IconSets.Item(18); $s = $ws.AutoFilter.Sort; $f = $s.SortFields.Add($ws.Range('E2:E21'), 3, 1); $f.SetIcon($wb.IconSets.Item(18).Item(1)); Apply $s }
Case 'table-icon-stars' 'table' { param($wb, $ws, $r, $lo) $c = $ws.Range('E2:E21').FormatConditions.Item(1); $c.IconSet = $wb.IconSets.Item(18); [void]$r.AutoFilter(5, $wb.IconSets.Item(18).Item(1), 10) }
Case 'table-sort-stars' 'table' { param($wb, $ws, $r, $lo) $c = $ws.Range('E2:E21').FormatConditions.Item(1); $c.IconSet = $wb.IconSets.Item(18); $s = $lo.Sort; $f = $s.SortFields.Add($ws.Range('E2:E21'), 3, 1); $f.SetIcon($wb.IconSets.Item(18).Item(1)); Apply $s }
Case 'table-pasted' 'table' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); $s = $lo.Sort; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); Apply $s; $lo.Range.Copy() | Out-Null; [void]$ws.Paste($ws.Range('H1')) }
Case 'table-moved-button' 'table' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, $m, 1, $m, $false); [void]$ws.Range('A1:F21').Cut($ws.Range('H3')) }

# a sheet's sort of a range outside a filter
Case 'range-sort' 'sheet' { param($wb, $ws, $r, $lo) $ws.AutoFilterMode = $false; $s = $ws.Sort; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); $s.SetRange($ws.Range('A1:F21')); Apply $s }
Case 'range-sort-case' 'sheet' { param($wb, $ws, $r, $lo) $ws.AutoFilterMode = $false; $s = $ws.Sort; $s.MatchCase = $true; [void]$s.SortFields.Add($ws.Range('B2:B21'), 0, 1); $s.SetRange($ws.Range('A1:F21')); Apply $s }
Case 'range-sort-columns' 'sheet' { param($wb, $ws, $r, $lo) $ws.AutoFilterMode = $false; $s = $ws.Sort; [void]$s.SortFields.Add($ws.Range('A3:F3'), 0, 1); $s.SetRange($ws.Range('A2:F21')); $s.Header = 2; $s.Orientation = 2; $s.Apply() }
Case 'range-sort-fill' 'sheet' { param($wb, $ws, $r, $lo) $ws.AutoFilterMode = $false; $s = $ws.Sort; $f = $s.SortFields.Add($ws.Range('C2:C21'), 1, 1); $f.SortOnValue.Color = $RED; $s.SetRange($ws.Range('A1:F21')); Apply $s }
Case 'range-sort-beside-filter' 'sheet' { param($wb, $ws, $r, $lo) $ws.Range('H1').Value2 = 'K'; $ws.Range('H2').Value2 = 2; $ws.Range('H3').Value2 = 1; $s = $ws.Sort; [void]$s.SortFields.Add($ws.Range('H2:H3'), 0, 1); $s.SetRange($ws.Range('H1:H3')); Apply $s }
Case 'range-sort-column-deleted' 'sheet' { param($wb, $ws, $r, $lo) $ws.AutoFilterMode = $false; $s = $ws.Sort; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); $s.SetRange($ws.Range('A1:F21')); Apply $s; [void]$ws.Columns.Item(3).Delete() }
$rangeSort = { param($ws) $ws.AutoFilterMode = $false; $s = $ws.Sort; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); $s.SetRange($ws.Range('A1:F21')); Apply $s }
Case 'range-sort-row-inside' 'sheet' { param($wb, $ws, $r, $lo) & $rangeSort $ws; [void]$ws.Rows.Item(5).Insert() }
Case 'range-sort-row-deleted' 'sheet' { param($wb, $ws, $r, $lo) & $rangeSort $ws; [void]$ws.Rows.Item(5).Delete() }
Case 'range-sort-column-before' 'sheet' { param($wb, $ws, $r, $lo) & $rangeSort $ws; [void]$ws.Columns.Item(2).Insert() }
Case 'range-sort-other-deleted' 'sheet' { param($wb, $ws, $r, $lo) & $rangeSort $ws; [void]$ws.Columns.Item(1).Delete() }
Case 'range-sort-moved' 'sheet' { param($wb, $ws, $r, $lo) & $rangeSort $ws; [void]$ws.Range('A1:F21').Cut($ws.Range('H3')) }
Case 'range-sort-moved-part' 'sheet' { param($wb, $ws, $r, $lo) & $rangeSort $ws; [void]$ws.Range('C2:C21').Cut($ws.Range('H3')) }
Case 'range-sort-moved-other' 'sheet' { param($wb, $ws, $r, $lo) & $rangeSort $ws; $o = $wb.Worksheets.Add(); $o.Name = 'Other'; [void]$ws.Range('A1:F21').Cut($o.Range('A1')) }
Case 'range-sort-sheet-copied' 'sheet' { param($wb, $ws, $r, $lo) & $rangeSort $ws; $ws.Copy($m, $ws) }
Case 'range-sort-table-over' 'sheet' { param($wb, $ws, $r, $lo) & $rangeSort $ws; [void]$ws.ListObjects.Add(1, $ws.Range('A1:F21'), $m, 1) }
Case 'range-sort-filter-set' 'sheet' { param($wb, $ws, $r, $lo) & $rangeSort $ws; [void]$ws.Range('A1:F21').AutoFilter() }

# a table's own forms
Case 'table-no-autofilter' 'table' { param($wb, $ws, $r, $lo) $lo.ShowAutoFilter = $false }
Case 'table-no-dropdown' 'table' { param($wb, $ws, $r, $lo) $lo.ShowAutoFilterDropDown = $false }
Case 'table-filter-then-off' 'table' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); $lo.ShowAutoFilter = $false }
Case 'table-filter-no-dropdown' 'table' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); $lo.ShowAutoFilterDropDown = $false }
Case 'table-no-headers' 'table' { param($wb, $ws, $r, $lo) $lo.ShowHeaders = $false }
Case 'table-no-headers-filter' 'table' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); $lo.ShowHeaders = $false }
Case 'table-totals-filter' 'table' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); $lo.ShowTotals = $true }
Case 'table-sort-range' 'table' { param($wb, $ws, $r, $lo) $s = $ws.Sort; [void]$s.SortFields.Add($ws.Range('C2:C21'), 0, 2); $s.SetRange($ws.Range('A1:F21')); Apply $s }
Case 'table-unlisted' 'table' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); $lo.Unlist() }
Case 'table-deleted-column-of-two' 'table' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(2, 'Banana'); [void]$r.AutoFilter(3, '>20'); $lo.ListColumns.Item('Product').Delete() }
Case 'table-column-added' 'table' { param($wb, $ws, $r, $lo) [void]$r.AutoFilter(3, '>20'); [void]$lo.ListColumns.Add(2) }
$xl.Quit()
$out | ConvertTo-Json -Depth 4
