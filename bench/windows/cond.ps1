# Conditional formats for xl.ps1 (bench/windows/cond.ts): rules made through the object model, what cells show by their
# display format (with their number format for bench/windows/lines.ts), and the rules each sheet lists. Dot-sourced by
# xl.ps1, whose Str, Cell and List it uses.

# a rule as bench/coverage/cond/cases.ts describes it (CondMake); one with a style gets a red fill and a dark red bold font
function CondMake($wb, $s) {
  $fcs = $wb.Worksheets.Item($s.sheet).Range($s.range).FormatConditions
  $m = [Type]::Missing
  $styled = $true
  switch ($s.add) {
    'cellIs' { $fc = if ($s.f2) { $fcs.Add(1, [int]$s.operator, [string]$s.f1, [string]$s.f2) } else { $fcs.Add(1, [int]$s.operator, [string]$s.f1) } }
    'expression' { $fc = $fcs.Add(2, $m, [string]$s.f1) }
    'text' { $fc = $fcs.Add(9, $m, $m, $m, [string]$s.text, [int]$s.operator) }
    'blanks' { $fc = $fcs.Add(10) }
    'noBlanks' { $fc = $fcs.Add(13) }
    'errors' { $fc = $fcs.Add(16) }
    'noErrors' { $fc = $fcs.Add(17) }
    'period' { $fc = $fcs.Add(11, $m, $m, $m, $m, $m, [int]$s.n) }
    'top10' {
      $fc = $fcs.AddTop10()
      $fc.TopBottom = $(if ($s.bottom) { 0 } else { 1 })
      $fc.Rank = [int]$s.n
      $fc.Percent = [bool]$s.percent
    }
    'aboveAverage' {
      $fc = $fcs.AddAboveAverage()
      $fc.AboveBelow = [int]$s.n
      if ($s.stdDev) { $fc.NumStdDev = [int]$s.stdDev }
    }
    'unique' { $fc = $fcs.AddUniqueValues(); $fc.DupeUnique = 0 }
    'duplicate' { $fc = $fcs.AddUniqueValues(); $fc.DupeUnique = 1 }
    'colorScale' { $fc = $fcs.AddColorScale([int]$s.n); $styled = $false }
    'dataBar' { $fc = $fcs.AddDatabar(); $styled = $false }
    'iconSet' { $fc = $fcs.AddIconSetCondition(); $fc.IconSet = $wb.IconSets.Item([int]$s.n); $styled = $false }
    default { throw "unknown rule $($s.add)" }
  }
  if ($styled) {
    # #FFC7CE and #9C0006 as COM's BGR
    $fc.Interior.Color = 13551615
    $fc.Font.Color = 393372
    $fc.Font.Bold = $true
  }
}

# the read cells: [sheet, [[address, Range.Text, the display format's fill (BGR), its colour index, font colour, bold, italic]]],
# with codes the cell's NumberFormat last
function CondLooks($wb, $reads, $codes) {
  List ($reads | ForEach-Object {
      $ws = $wb.Worksheets.Item([string]$_[0])
      $cells = foreach ($c in $ws.Range([string]$_[1]).Cells) {
        $d = $c.DisplayFormat
        $look = @((Str $c.Address($false, $false)), (Str ([string]$c.Text)), (Cell ([double]$d.Interior.Color)), (Cell ([double]$d.Interior.ColorIndex)), (Cell ([double]$d.Font.Color)), (Cell ([bool]$d.Font.Bold)), (Cell ([bool]$d.Font.Italic)))
        if ($codes) { $look += @(Str ([string]$c.NumberFormat)) }
        List $look
      }
      List @((Str $ws.Name), (List $cells))
    })
}

# each sheet's rules as the object model lists them: [sheet, [[type, applies to, formula 1, priority, stop if true]]]
function CondRules($wb) {
  List ($wb.Worksheets | ForEach-Object {
      $ws = $_
      $rules = foreach ($fc in $ws.Cells.FormatConditions) {
        $f = try { [string]$fc.Formula1 } catch { '' }
        $stop = try { [bool]$fc.StopIfTrue } catch { $false }
        List @((Cell ([double]$fc.Type)), (Str $fc.AppliesTo.Address($false, $false)), (Str $f), (Cell ([double]$fc.Priority)), (Cell $stop))
      }
      List @((Str $ws.Name), (List $rules))
    })
}

function Cond($wb, $s) {
  switch ($s.kind) {
    'make' { try { CondMake $wb $s; 'null' } catch { Err $_ } }
    'looks' { CondLooks $wb $s.read $s.codes }
    'rules' { CondRules $wb }
    default { throw "unknown cond step $($s.kind)" }
  }
}
