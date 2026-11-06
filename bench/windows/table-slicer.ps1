# Slicers of tables as Excel for Windows makes and saves them, one book a case, and what it holds after: the forms
# [MS-XLSX] does not give (the extensions of the workbook and the sheet, the frame) and what a change of the table does.
param([string]$dir)
$ErrorActionPreference = 'Continue'
New-Item -ItemType Directory -Force -Path $dir | Out-Null
$xl = New-Object -ComObject Excel.Application
$xl.DisplayAlerts = $false
$xl.Visible = $false
$m = [Type]::Missing
$out = @()
# a table Sales of Region, Product, Amount with a slicer of Product; with $gap a first column deleted, so the others' ids are 2, 3, 4
function Case($id, $gap, [scriptblock]$act) {
  $wb = $xl.Workbooks.Add()
  while ($wb.Worksheets.Count -gt 1) { $wb.Worksheets.Item($wb.Worksheets.Count).Delete() }
  $ws = $wb.Worksheets.Item(1)
  $ws.Name = 'Data'
  $rows = @(@('Gone','Region','Product','Amount'),@('x','North','Apple',10),@('x','South','Banana',20),@('x','East','Apple',5),@('x','North','Cherry',7),@('x','West','Banana',1))
  for ($r = 0; $r -lt $rows.Count; $r++) { for ($c = 0; $c -lt 4; $c++) { $ws.Range(('ABCD'[$c]).ToString() + ($r + 1)).Formula = [string]$rows[$r][$c] } }
  if (-not $gap) { [void]$ws.Columns.Item(1).Delete() }
  $lo = $ws.ListObjects.Add(1, $ws.Range($(if ($gap) { 'A1:D6' } else { 'A1:C6' })), $m, 1)
  $lo.Name = 'Sales'
  if ($gap) { $lo.ListColumns.Item('Gone').Delete() }
  $sc = $wb.SlicerCaches.Add2($lo, 'Product')
  $sl = $sc.Slicers.Add($ws, $m, 'Product', 'Product', 10, 300, 144, 198)
  $err = ''
  try { & $act $wb $ws $lo $sc $sl } catch { $err = $_.Exception.Message }
  try { $wb.SaveAs("$dir\$id.xlsx", 51) } catch { $err += ' SAVE ' + $_.Exception.Message }
  $names = @(); foreach ($n in $wb.Names) { $names += [string]$n.Name }
  $caches = @(); foreach ($c in $wb.SlicerCaches) { $caches += @{ name = [string]$c.Name; source = [string]$c.SourceName; slicers = [int]$c.Slicers.Count } }
  $script:out += @{ id = $id; refused = $err; names = $names; caches = $caches }
  $wb.Close($false)
}
Case 'plain' $true { param($wb,$ws,$lo,$sc,$sl) }
Case 'chosen' $true { param($wb,$ws,$lo,$sc,$sl) $sc.SlicerItems.Item('Banana').Selected = $false }
Case 'sort-desc' $false { param($wb,$ws,$lo,$sc,$sl) $sc.SortItems = 3 }
Case 'sort-source' $false { param($wb,$ws,$lo,$sc,$sl) $sc.SortItems = 1 }
Case 'no-custom' $false { param($wb,$ws,$lo,$sc,$sl) $sc.SortUsingCustomLists = $false }
Case 'cross-none' $false { param($wb,$ws,$lo,$sc,$sl) $sc.CrossFilterType = 1 }
Case 'cross-last' $false { param($wb,$ws,$lo,$sc,$sl) $sc.CrossFilterType = 3 }
Case 'hide-nodata' $false { param($wb,$ws,$lo,$sc,$sl) $sc.ShowAllItems = $false }
Case 'column-renamed' $false { param($wb,$ws,$lo,$sc,$sl) $ws.Range('B1').Value2 = 'Goods' }
Case 'column-deleted' $false { param($wb,$ws,$lo,$sc,$sl) $lo.ListColumns.Item('Product').Delete() }
Case 'other-column-deleted' $false { param($wb,$ws,$lo,$sc,$sl) $lo.ListColumns.Item('Region').Delete() }
Case 'table-renamed' $false { param($wb,$ws,$lo,$sc,$sl) $lo.Name = 'Goods' }
Case 'table-deleted' $false { param($wb,$ws,$lo,$sc,$sl) $lo.Delete() }
Case 'table-unlisted' $false { param($wb,$ws,$lo,$sc,$sl) $lo.Unlist() }
Case 'two-slicers' $false { param($wb,$ws,$lo,$sc,$sl) [void]$sc.Slicers.Add($ws, $m, 'Product 1', 'Other', 10, 500, 144, 198) }
Case 'second-column' $false { param($wb,$ws,$lo,$sc,$sl) $c2 = $wb.SlicerCaches.Add2($lo, 'Region'); [void]$c2.Slicers.Add($ws, $m, 'Region', 'Region', 10, 500, 144, 198) }
Case 'no-headers' $false { param($wb,$ws,$lo,$sc,$sl) $lo.ShowHeaders = $false }
Case 'no-autofilter' $false { param($wb,$ws,$lo,$sc,$sl) $lo.ShowAutoFilter = $false }
Case 'slicer-deleted' $false { param($wb,$ws,$lo,$sc,$sl) $sl.Delete() }
$xl.Quit()
$out | ConvertTo-Json -Depth 4
