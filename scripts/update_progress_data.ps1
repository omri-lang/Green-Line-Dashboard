param(
  [Parameter(Mandatory=$true)][string]$InFolder,
  [Parameter(Mandatory=$true)][string]$OutJs
)

function Parse-DateFromName($name) {
  # Match patterns like 28.06.2026 or 3.6.26 etc, keep the dd.mm.yyyy style
  if ($name -match '(\d{1,2})\.(\d{1,2})\.(\d{4})') {
    $d = [int]$matches[1]; $m = [int]$matches[2]; $y = [int]$matches[3]
    return (Get-Date -Year $y -Month $m -Day $d).ToString('yyyy-MM-dd')
  }
  return $null
}

# The 8 disciplines shown on the progress map. Matched case-insensitively against
# the trimmed discipline-block name (row 3), which sits 4 columns before each
# "Execution from Cumulative (%)" column (row 5).
$TARGET_DISCIPLINES = @(
  "Excavation to foundation",
  "Bedding Refill",
  "OHLE piles",
  "Multitubular",
  "Foundations slab",
  "Track Installation on line",
  "TRK Turnout Installation",
  "OHLE Poles Installation"
)

# Section 7A is a bored tunnel, reported on its own "Tunnels 7A" sheet with a
# different (smaller) set of disciplines and its own raw column names. Map
# each raw sheet name to the canonical discipline key used above, so 7A's
# data merges into the exact same disciplines map / UI as every other section.
$TUNNEL_DISCIPLINE_MAP = @{
  "Foundation slab"             = "Foundations slab"
  "Multitabular"                 = "Multitubular"
  "Track Installation on line"  = "Track Installation on line"
}

function Find-DisciplineBlocks($ws, $maxCol, $nameLookup) {
  # nameLookup: hashtable of rawName -> canonicalName to match against row 3
  # (case-insensitive, trimmed). Returns canonicalName -> {unitsCol, actCol, planCol}.
  $blocks = @{}
  foreach ($c in 5..$maxCol) {
    $h5 = $ws.Cells.Item(5, $c).Value2
    if (-not ($h5 -is [string])) { continue }
    if ($h5.Trim() -ne "Execution from Cumulative (%)") { continue }
    $rawName = $ws.Cells.Item(3, $c - 4).Value2
    if (-not ($rawName -is [string])) { continue }
    $name = $rawName.Trim()
    foreach ($raw in $nameLookup.Keys) {
      if ($name -ieq $raw) {
        $canonical = $nameLookup[$raw]
        if (-not $blocks.ContainsKey($canonical)) {
          $blocks[$canonical] = @{ unitsCol = $c - 3; actCol = $c - 2; planCol = $c - 1 }
        }
      }
    }
  }
  return $blocks
}

function Add-SheetData($ws, $discBlocks, $lotLength, $lotDiscAgg, $lengthMode) {
  # lengthMode "sum": rows are distinct physical sub-segments, lengths add up (Main line).
  # lengthMode "max": rows are parallel bores/structures of the same section length,
  # so length should not be double-counted (Tunnels 7A: 2 tunnel tubes, same length).
  $used = $ws.UsedRange
  $maxRow = $used.Row + $used.Rows.Count - 1

  for ($r = 6; $r -le $maxRow; $r++) {
    $gsec = $ws.Cells.Item($r, 2).Value2
    $loc  = $ws.Cells.Item($r, 3).Value2
    $len  = $ws.Cells.Item($r, 4).Value2
    if ([string]::IsNullOrWhiteSpace([string]$gsec) -and [string]::IsNullOrWhiteSpace([string]$loc)) { continue }
    if (-not ($loc -is [string]) -or [string]::IsNullOrWhiteSpace($loc)) { continue }
    if (-not ($len -is [double]) -or $len -le 0) { continue }

    if (-not $lotLength.ContainsKey($loc)) { $lotLength[$loc] = 0.0 }
    if ($lengthMode -eq "max") {
      if ($len -gt $lotLength[$loc]) { $lotLength[$loc] = $len }
    } else {
      $lotLength[$loc] += $len
    }

    if (-not $lotDiscAgg.ContainsKey($loc)) { $lotDiscAgg[$loc] = @{} }

    foreach ($target in $discBlocks.Keys) {
      $cols = $discBlocks[$target]
      $act  = $ws.Cells.Item($r, $cols.actCol).Value2
      $plan = $ws.Cells.Item($r, $cols.planCol).Value2
      $unit = $ws.Cells.Item($r, $cols.unitsCol).Value2
      # A blank/non-numeric "Total (Planned)" means this discipline doesn't apply
      # to this row at all - skip it. A blank "Act. Cumulative" with a real planned
      # quantity means genuine zero progress so far (not "doesn't apply") - count it.
      if (-not ($plan -is [double]) -or $plan -le 0) { continue }
      $actVal = if ($act -is [double]) { $act } else { 0.0 }

      if (-not $lotDiscAgg[$loc].ContainsKey($target)) {
        $lotDiscAgg[$loc][$target] = @{ sumActual = 0.0; sumPlanned = 0.0; unit = $null }
      }
      $entry = $lotDiscAgg[$loc][$target]
      $entry.sumActual += $actVal
      $entry.sumPlanned += $plan
      if (-not $entry.unit -and ($unit -is [string]) -and -not [string]::IsNullOrWhiteSpace($unit)) {
        $entry.unit = $unit.Trim()
      }
      $lotDiscAgg[$loc][$target] = $entry
    }
  }
}

# Depot "-Buildings" sheets: one row per building, column A = that building's
# weight/portion within the depot's total structure scope (0-1, sums to ~1
# across all buildings), column B = building name, column D = that building's
# own overall Execution % (used as a per-building "status" reference shown
# under every discipline row - it's the same figure regardless of which
# discipline you're looking at). Each of the first 3 disciplines lives in its
# own block whose row-3 header gives the label, with "Execution %" exactly
# one column to the right of the label column. "RSY - General" is laid out
# differently: its label sits at row 4 in the SAME column as its own
# "Execution %" (row 5) - verified directly against both Holon (col 76/BX)
# and Herz (col 65/BM) on the newest report.
$BUILDING_DISCIPLINE_LABELS = [ordered]@{
  "Excavation and Foundation" = @("Excavation and Foundation")
  "Shell"                     = @("Skeleton Works", "Shell")
  "Arch and MEP Works"        = @("Arch and MEP Works")
}
$RSY_GENERAL_LABEL = "RSY - General"

# Building rows appear in a fixed order per sheet; two Holon rows share the
# literal name "Storm Water Reservoire", so buildings are matched by position
# (row order), not by name lookup.
$HOLON_BUILDING_NAMES = @("מבנה הנהלה","מבנה תחזוקה יומי","מבנה נהגים","מבנה תחזוקה כבדה","מבנה כניסה","TTR-01","בור ניקוז 1","בור ניקוז 2")
$HERZ_BUILDING_NAMES  = @("מבנה הנהלה","מבנה תחזוקה יומי","מבנה ניקיון","מבנה תחזוקה קלה","מבנה תחזוקה","מבנה כניסה","TTR-15","בור ניקוז")

# Depot "-Yard" sheets: a small rollup table (header row has "Activity" in
# column D) whose row position varies by sheet - locate it by searching
# column D rather than assuming a fixed row. Columns D-G = code / actual /
# planned / pct.
$YARD_DEPOT_MAP = [ordered]@{
  "MT"          = "Multitubular"
  "FS"          = "Foundations slab"
  "Track"       = "Track Installation on line"
  "OHLE poles"  = "OHLE Poles Installation"
}

function Get-BuildingsData($ws, $buildingNames) {
  $used = $ws.UsedRange
  $maxRow = $used.Row + $used.Rows.Count - 1
  $maxCol = $used.Column + $used.Columns.Count - 1

  $discCols = @{}
  foreach ($canonical in $BUILDING_DISCIPLINE_LABELS.Keys) {
    foreach ($lbl in $BUILDING_DISCIPLINE_LABELS[$canonical]) {
      for ($c = 1; $c -le $maxCol; $c++) {
        $r3 = $ws.Cells.Item(3, $c).Value2
        if ($r3 -is [string] -and $r3.Trim() -ieq $lbl) { $discCols[$canonical] = $c + 1; break }
      }
      if ($discCols.ContainsKey($canonical)) { break }
    }
  }
  for ($c = 1; $c -le $maxCol; $c++) {
    $r4 = $ws.Cells.Item(4, $c).Value2
    if ($r4 -is [string] -and $r4.Trim() -ieq $RSY_GENERAL_LABEL) { $discCols[$RSY_GENERAL_LABEL] = $c; break }
  }

  $sumPct = @{}
  $countBldg = @{}
  $buildingLists = @{}
  foreach ($k in $discCols.Keys) { $sumPct[$k] = 0.0; $countBldg[$k] = 0; $buildingLists[$k] = @() }

  $bIndex = 0
  for ($r = 6; $r -le $maxRow; $r++) {
    $name = $ws.Cells.Item($r, 2).Value2
    if (-not ($name -is [string]) -or [string]::IsNullOrWhiteSpace($name)) { continue }
    if ($name.Trim() -eq "Weighted Average") { continue }
    $heName = if ($bIndex -lt $buildingNames.Count) { $buildingNames[$bIndex] } else { $name.Trim() }

    # Per-building percentage must come from THIS discipline's own column, not
    # column D (that building's overall % across all disciplines) - otherwise
    # a discipline whose average is 100% could show buildings well under 100%
    # underneath it, which is exactly the bug reported: reusing one shared
    # column-D list under every discipline row instead of each discipline's
    # own per-building figures.
    #
    # Flat, equally-weighted average across buildings (each building counts
    # once, regardless of its budget/weight share in column A) - per Omri's
    # request, replacing the previous weight-share weighted average.
    foreach ($canonical in $discCols.Keys) {
      $pct = $ws.Cells.Item($r, $discCols[$canonical]).Value2
      if (-not ($pct -is [double])) { continue }
      $sumPct[$canonical] += $pct
      $countBldg[$canonical] += 1
      $buildingLists[$canonical] += [ordered]@{ name = $heName; pct = [math]::Round(100.0 * $pct, 1) }
    }
    $bIndex++
  }

  $disciplines = [ordered]@{}
  foreach ($canonical in (@($BUILDING_DISCIPLINE_LABELS.Keys) + @($RSY_GENERAL_LABEL))) {
    if (-not $discCols.ContainsKey($canonical)) { continue }
    if ($countBldg[$canonical] -le 0) { continue }
    $disciplines[$canonical] = [ordered]@{
      pct         = [math]::Round(100.0 * $sumPct[$canonical] / $countBldg[$canonical], 1)
      buildingList = $buildingLists[$canonical]
    }
  }
  return [ordered]@{ disciplines = $disciplines }
}

function Get-DepotYardData($ws) {
  $used = $ws.UsedRange
  $maxRow = $used.Row + $used.Rows.Count - 1

  $headerRow = $null
  for ($r = 1; $r -le $maxRow; $r++) {
    $d = $ws.Cells.Item($r, 4).Value2
    if ($d -is [string] -and $d.Trim() -eq "Activity") { $headerRow = $r; break }
  }
  if (-not $headerRow) { return [ordered]@{} }

  $result = [ordered]@{}
  for ($r = $headerRow + 1; $r -le $headerRow + 6; $r++) {
    $code = $ws.Cells.Item($r, 4).Value2
    if (-not ($code -is [string])) { continue }
    $code = $code.Trim()
    if (-not $YARD_DEPOT_MAP.Contains($code)) { continue }
    $canonical = $YARD_DEPOT_MAP[$code]
    $act = $ws.Cells.Item($r, 5).Value2
    $plan = $ws.Cells.Item($r, 6).Value2
    if (-not ($plan -is [double]) -or $plan -le 0) { continue }
    $actVal = if ($act -is [double]) { $act } else { 0.0 }
    $result[$canonical] = [ordered]@{
      pct     = [math]::Round(100.0 * $actVal / $plan, 1)
      actual  = [math]::Round($actVal, 1)
      planned = [math]::Round($plan, 1)
    }
  }
  return $result
}

function Add-JunctionData($ws, $lotJunctions) {
  # "Junctions" sheet: one row per junction. Column B ("A section") is the
  # 1A-12A section code, column D ("Name") is the junction's English name
  # (kept verbatim, just trimmed), column G ("Execution %") is that
  # junction's own completion fraction (0-1). A trailing "Average" summary
  # row and blank rows have an empty column B and are skipped naturally.
  # "6A-2" is not a sub-code of 6A - it denotes section 7A's junctions and
  # must map there.
  $used = $ws.UsedRange
  $maxRow = $used.Row + $used.Rows.Count - 1

  for ($r = 5; $r -le $maxRow; $r++) {
    $loc = $ws.Cells.Item($r, 2).Value2
    if (-not ($loc -is [string]) -or [string]::IsNullOrWhiteSpace($loc)) { continue }
    $loc = $loc.Trim()
    if ($loc -eq "6A-2") { $loc = "7A" }
    elseif ($loc -match '^(\d+A)-\d+$') { $loc = $matches[1] }

    $pct = $ws.Cells.Item($r, 7).Value2
    if (-not ($pct -is [double])) { continue }

    $name = $ws.Cells.Item($r, 4).Value2
    $name = if ($name -is [string]) { $name.Trim() } else { "" }

    if (-not $lotJunctions.ContainsKey($loc)) {
      $lotJunctions[$loc] = @{ total = 0; completed = 0; inProgress = 0; completedNames = @(); inProgressNames = @() }
    }
    $lotJunctions[$loc].total++
    if ($pct -ge 1) {
      $lotJunctions[$loc].completed++
      if ($name) { $lotJunctions[$loc].completedNames += $name }
    }
    elseif ($pct -gt 0) {
      $lotJunctions[$loc].inProgress++
      if ($name) { $lotJunctions[$loc].inProgressNames += $name }
    }
  }
}

$DEPOT_SHEETS = [ordered]@{
  "Holon" = @{ buildings = "Depot Holon-Buildings"; yard = "Depot Holon-Yard" }
  "Herz"  = @{ buildings = "Depot Herz-Buildings";  yard = "Depot Herz-Yard" }
}

# "TTR" sheet: one row per turnout structure. Column B = structure name (kept
# verbatim, e.g. "TTR2"), column E = overall "Execution %" (row 5 header),
# columns O/AA/AJ/AS/BC = the same 5 sub-disciplines' own "Execution %"
# (row 5 header, row 4 gives each block's label) - verified identical layout
# on both the oldest (29.1.2026) and newest (05.07.2026) reports.
$TTR_DETAIL_COLS = [ordered]@{
  "Building Foundation"         = 15  # O
  "Shell"                       = 27  # AA
  "MEP and Architectural Works" = 36  # AJ
  "Rail Systems"                = 45  # AS
  "Landscaping and Completion"  = 55  # BC
}

function Get-TTRData($ws) {
  $used = $ws.UsedRange
  $maxRow = $used.Row + $used.Rows.Count - 1
  $list = @()
  for ($r = 6; $r -le $maxRow; $r++) {
    $name = $ws.Cells.Item($r, 2).Value2
    if (-not ($name -is [string]) -or [string]::IsNullOrWhiteSpace($name)) { continue }

    $details = [ordered]@{}
    foreach ($key in $TTR_DETAIL_COLS.Keys) {
      $v = $ws.Cells.Item($r, $TTR_DETAIL_COLS[$key]).Value2
      if ($v -is [double]) { $details[$key] = [math]::Round(100.0 * $v, 1) }
    }
    if ($details.Count -eq 0) { continue }

    # Flat, equally-weighted average across the 5 discipline percentages (each
    # weighted 1/5) - per Omri's request, replacing Excel's own column-E
    # "Execution %" (a budget-weighted formula) as this row's headline %.
    $sum = 0.0
    foreach ($v in $details.Values) { $sum += $v }
    $avgPct = [math]::Round($sum / $details.Count, 1)

    $list += [ordered]@{
      name    = $name.Trim()
      pct     = $avgPct
      details = $details
    }
  }
  return $list
}

$files = Get-ChildItem -Path $InFolder -Filter "Construction Bi-Weekly report-*.xlsx" | Sort-Object Name
$results = @{}
$depotResults = @{}
$ttrResults = @{}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

foreach ($f in $files) {
  $dateStr = Parse-DateFromName $f.Name
  if (-not $dateStr) { Write-Output "SKIP (no date): $($f.Name)"; continue }

  try {
    $wb = $excel.Workbooks.Open($f.FullName, 0, $true)
    $sheetName = $null
    foreach ($cand in @("Main line", "AG")) {
      foreach ($s in $wb.Worksheets) { if ($s.Name -eq $cand) { $sheetName = $cand; break } }
      if ($sheetName) { break }
    }
    if (-not $sheetName) { throw "No Main line / AG sheet found" }
    $ws = $wb.Worksheets.Item($sheetName)
    $used = $ws.UsedRange
    $maxCol = $used.Column + $used.Columns.Count - 1

    # Main line / AG: real per-discipline block match is the exact header
    # "Execution from Cumulative (%)" (not the bare substring "Execution", which
    # also matches unrelated summary columns like "Execution %" on "* - General"
    # blocks and produced an inflated/incorrect composite in the old version of
    # this script).
    $mainLookup = @{}
    foreach ($t in $TARGET_DISCIPLINES) { $mainLookup[$t] = $t }
    $discBlocks = Find-DisciplineBlocks $ws $maxCol $mainLookup
    foreach ($target in $TARGET_DISCIPLINES) {
      if (-not $discBlocks.ContainsKey($target)) {
        Write-Output "WARN $($f.Name): discipline not found on sheet '$sheetName': $target"
      }
    }

    $lotLength = @{}
    $lotDiscAgg = @{}
    Add-SheetData $ws $discBlocks $lotLength $lotDiscAgg "sum"

    # Section 7A (bored tunnel) lives on its own "Tunnels 7A" sheet with just 3
    # disciplines, present only in newer reports. Merge it into the same
    # per-lot dictionaries so 7A comes out through the exact same code path.
    $tunnelSheet = $null
    $tunnelBlocks = @{}
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq "Tunnels 7A") { $tunnelSheet = $s.Name } }
    if ($tunnelSheet) {
      $wsT = $wb.Worksheets.Item($tunnelSheet)
      $usedT = $wsT.UsedRange
      $maxColT = $usedT.Column + $usedT.Columns.Count - 1
      $tunnelBlocks = Find-DisciplineBlocks $wsT $maxColT $TUNNEL_DISCIPLINE_MAP
      foreach ($raw in $TUNNEL_DISCIPLINE_MAP.Keys) {
        $canonical = $TUNNEL_DISCIPLINE_MAP[$raw]
        if (-not $tunnelBlocks.ContainsKey($canonical)) {
          Write-Output "WARN $($f.Name): tunnel discipline not found on sheet 'Tunnels 7A': $raw"
        }
      }
      Add-SheetData $wsT $tunnelBlocks $lotLength $lotDiscAgg "max"
    }

    # Junctions: one completed/in-progress count per section, from its own sheet.
    $lotJunctions = @{}
    $junctionsSheet = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq "Junctions") { $junctionsSheet = $s.Name } }
    if ($junctionsSheet) {
      Add-JunctionData $wb.Worksheets.Item($junctionsSheet) $lotJunctions
    }

    # Build final per-lot shape: overall (quantity-weighted across whichever
    # disciplines apply to that section, so it's always consistent with what
    # the breakdown panel shows) + disciplines map (only entries where planned > 0).
    $lotFinal = @{}
    $allLocs = @{}
    foreach ($k in $lotDiscAgg.Keys) { $allLocs[$k] = $true }
    foreach ($k in $lotJunctions.Keys) { $allLocs[$k] = $true }

    foreach ($loc in $allLocs.Keys) {
      $disciplines = [ordered]@{}
      $sumActualAll = 0.0
      $sumPlannedAll = 0.0

      if ($lotDiscAgg.ContainsKey($loc)) {
        foreach ($target in $TARGET_DISCIPLINES) {
          if (-not $lotDiscAgg[$loc].ContainsKey($target)) { continue }
          $entry = $lotDiscAgg[$loc][$target]
          if ($entry.sumPlanned -le 0) { continue }
          $disciplines[$target] = [ordered]@{
            pct     = [math]::Round(100.0 * $entry.sumActual / $entry.sumPlanned, 1)
            actual  = [math]::Round($entry.sumActual, 1)
            planned = [math]::Round($entry.sumPlanned, 1)
            unit    = $(if ($entry.unit) { $entry.unit } else { "" })
          }
          $sumActualAll += $entry.sumActual
          $sumPlannedAll += $entry.sumPlanned
        }
      }

      if ($disciplines.Count -eq 0 -and -not $lotJunctions.ContainsKey($loc)) { continue }

      $lotObj = [ordered]@{}
      $lotObj.overall = [ordered]@{
        pct = $(if ($sumPlannedAll -gt 0) { [math]::Round(100.0 * $sumActualAll / $sumPlannedAll, 1) } else { 0 })
        len = [math]::Round($(if ($lotLength.ContainsKey($loc)) { $lotLength[$loc] } else { 0 }), 0)
      }
      $lotObj.disciplines = $disciplines
      if ($lotJunctions.ContainsKey($loc)) {
        $lotObj.junctions = [ordered]@{
          total           = $lotJunctions[$loc].total
          completed       = $lotJunctions[$loc].completed
          inProgress      = $lotJunctions[$loc].inProgress
          completedNames  = $lotJunctions[$loc].completedNames
          inProgressNames = $lotJunctions[$loc].inProgressNames
        }
      }
      $lotFinal[$loc] = $lotObj
    }

    $results[$dateStr] = $lotFinal

    # Depots (Holon / Herz): 4 building disciplines (weighted-avg % across
    # buildings) + per-building overall-% list + 4 yard disciplines (real
    # actual/planned quantities).
    $depotsForDate = [ordered]@{}
    foreach ($depotKey in $DEPOT_SHEETS.Keys) {
      $sheetNames = $DEPOT_SHEETS[$depotKey]
      $bldSheet = $null; $yardSheet = $null
      foreach ($s in $wb.Worksheets) {
        if ($s.Name -eq $sheetNames.buildings) { $bldSheet = $s.Name }
        if ($s.Name -eq $sheetNames.yard) { $yardSheet = $s.Name }
      }
      if (-not $bldSheet -and -not $yardSheet) { continue }
      $depotObj = [ordered]@{}
      if ($bldSheet) {
        $buildingNames = if ($depotKey -eq "Holon") { $HOLON_BUILDING_NAMES } else { $HERZ_BUILDING_NAMES }
        $bd = Get-BuildingsData $wb.Worksheets.Item($bldSheet) $buildingNames
        $depotObj.buildings = $bd.disciplines
      }
      if ($yardSheet) { $depotObj.yard = Get-DepotYardData $wb.Worksheets.Item($yardSheet) }
      $depotsForDate[$depotKey] = $depotObj
    }
    $depotResults[$dateStr] = $depotsForDate

    # TTR: one row per turnout structure, own name + overall % + 5 sub-discipline %s.
    $ttrSheet = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq "TTR") { $ttrSheet = $s.Name } }
    $ttrList = @()
    if ($ttrSheet) { $ttrList = Get-TTRData $wb.Worksheets.Item($ttrSheet) }
    $ttrResults[$dateStr] = $ttrList

    Write-Output "OK $dateStr <- $($f.Name)  lots=$($lotFinal.Count) disciplines=$($discBlocks.Count)/$($TARGET_DISCIPLINES.Count) tunnel=$($tunnelBlocks.Count -as [string]) junctionSections=$($lotJunctions.Count) depots=$($depotsForDate.Count) ttr=$($ttrList.Count)"

    $wb.Close($false)
  } catch {
    Write-Output "ERROR $($f.Name): $($_.Exception.Message)"
    try { $wb.Close($false) } catch {}
  }
}

$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

# build ordered JSON
$sortedDates = $results.Keys | Sort-Object
$obj = [ordered]@{}
foreach ($d in $sortedDates) {
  $lotsOrdered = [ordered]@{}
  foreach ($k in ($results[$d].Keys | Sort-Object)) { $lotsOrdered[$k] = $results[$d][$k] }
  $obj[$d] = $lotsOrdered
}
$json = $obj | ConvertTo-Json -Depth 6

$depotSortedDates = $depotResults.Keys | Sort-Object
$depotObjAll = [ordered]@{}
foreach ($d in $depotSortedDates) { $depotObjAll[$d] = $depotResults[$d] }
$depotJson = $depotObjAll | ConvertTo-Json -Depth 6

$ttrSortedDates = $ttrResults.Keys | Sort-Object
$ttrObjAll = [ordered]@{}
foreach ($d in $ttrSortedDates) { $ttrObjAll[$d] = $ttrResults[$d] }
$ttrJson = $ttrObjAll | ConvertTo-Json -Depth 6

$js = "// Auto-generated by update_progress_data.ps1 - do not edit by hand`r`nconst PROGRESS_DATA = $json;`r`nconst DEPOT_DATA = $depotJson;`r`nconst TTR_DATA = $ttrJson;`r`n"
Set-Content -Path $OutJs -Value $js -Encoding utf8
Write-Output "WROTE: $OutJs"
