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

# "CW Track surfacing" (ריצוף) and "OHLE cable" pulling/adjustment (כבלי חשמל) -
# same block layout as the 8 above (found via the same "Execution from
# Cumulative (%)" header match), but per Omri's request (2026-07-19) these are
# shown only as extra rows in each section's detail panel, not folded into the
# section's own headline % (map color / table "ביצוע" column), so they're kept
# in their own list rather than added to $TARGET_DISCIPLINES.
# "OHLE cable" is the raw row-3 label text (trimmed) for the "OHLE cable
# pulling and adjustment (5)" block (row 4 sub-header) - Find-DisciplineBlocks
# matches on the row-3 label, not row 4.
$EXTRA_DISCIPLINES = @(
  "CW Track surfacing",
  "OHLE cable"
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

# "Stops" sheet: one row per station. Row 3 gives each column-group's raw
# label ("Location", "Name", "Stops - General", "Secondary Routing and
# Manholes", "CSLM Casting", "SIG Stops equipment installation", "COM
# Equipment installation & Testing"); row 4 gives that same column's
# sub-header ("Execution %" for the columns we want). The exact column
# letters SHIFT between report versions (verified: N/AB/BD/CT on the
# 29.1.2026 report vs O/AC/BE/CU from 07.06.2026 onward), so columns must be
# located by their row-3 label per file rather than assumed fixed.
#
# The station list itself (Hebrew name + section) is not on this sheet at
# all - it's the canonical list from "טבלת תחנות.xlsx" (62 stations, order
# below follows route direction 1A->12A). Per Omri's instruction (2026-07-16):
# this Hebrew/section list is authoritative; any station the Stops sheet
# doesn't have under a matching English name is still shown, at 0%.
$STOPS_GROUP_LABELS = [ordered]@{
  location   = "Location"
  name       = "Name"
  general    = "Stops - General"
  secondary  = "Secondary Routing and Manholes"
  cslm       = "CSLM Casting"
  sig        = "SIG Stops equipment installation"
  com        = "COM Equipment installation & Testing"
  surfacing  = "CW Station - Platform Surfacing"
}

$STOPS_MASTER_LIST = @(
  @{ en = "Holon East"; he = "חולון מזרח"; section = "1A" }
  @{ en = "Agricultural Center"; he = "הקריה החקלאית"; section = "1A" }
  @{ en = "Peres Park"; he = "פארק פרס"; section = "2A" }
  @{ en = "HaMerkava"; he = "המרכבה"; section = "2A" }
  @{ en = "Holon Junction"; he = "צומת חולון"; section = "3A" }
  @{ en = "Holon Theater"; he = "תיאטרון חולון"; section = "3A" }
  @{ en = "Kugel"; he = "קוגל"; section = "3A" }
  @{ en = "Krauze"; he = "קראוזה"; section = "3A" }
  @{ en = "Sokolov East"; he = "סוקולוב מזרח"; section = "3A" }
  @{ en = "Ge'ulim"; he = "גאולים"; section = "3A" }
  @{ en = "Cactus Garden"; he = "גן הקקטוסים"; section = "3A" }
  @{ en = "HaMelakha"; he = "המלאכה"; section = "3A" }
  @{ en = "Nahalat Yehuda"; he = "נחלת יהודה"; section = "4A" }
  @{ en = "Sakharov"; he = "סחרוב"; section = "4A" }
  @{ en = "Yalde Tehran"; he = "ילדי טהרן"; section = "4A" }
  @{ en = "Lishansky"; he = "לישנסקי"; section = "4A" }
  @{ en = "Moshe Dayan"; he = "משה דיין"; section = "4A" }
  @{ en = "HaHistadrut"; he = "ההסתדרות"; section = "5A" }
  @{ en = "Holon Institute of Technology"; he = "המכון הטכנולוגי חולון"; section = "5A" }
  @{ en = "Lavon"; he = "לבון"; section = "5A" }
  @{ en = "Begin"; he = "בגין"; section = "5A" }
  @{ en = "Bar-Lev"; he = "בר-לב"; section = "5A" }
  @{ en = "Laskov"; he = "לסקוב"; section = "5A" }
  @{ en = "HaBonim Park"; he = "פארק הבונים"; section = "5A" }
  @{ en = "Shapira"; he = "שפירא"; section = "6A" }
  @{ en = "Kibbuts Galuyot"; he = "קיבוץ גלויות"; section = "6A" }
  @{ en = "HaHurshot Park"; he = "פארק החורשות"; section = "6A" }
  @{ en = "Kiryat Shalom"; he = "קריית שלום"; section = "6A" }
  @{ en = "Yehuda HaMakkabbi"; he = "יהודה המכבי"; section = "7A" }
  @{ en = "Ibn Gabirol Arlosoroff"; he = "אבן גבירול ארלוזורוב"; section = "7A" }
  @{ en = "Rabin Square"; he = "כיכר רבין"; section = "7A" }
  @{ en = "Kaplan"; he = "קפלן"; section = "7A" }
  @{ en = "Carlebach"; he = "קרליבך"; section = "7A" }
  @{ en = "Levinski Garden"; he = "גינת לוינסקי"; section = "7A" }
  @{ en = "Namir"; he = "נמיר"; section = "8A" }
  @{ en = "Brodetsky"; he = "ברודצקי"; section = "8A" }
  @{ en = "Tel Aviv University"; he = "אוניברסיטת תל אביב"; section = "8A" }
  @{ en = "Broshim"; he = "ברושים"; section = "8A" }
  @{ en = "Reading"; he = "רדינג"; section = "8A" }
  @{ en = "HaYeridim"; he = "הירידים"; section = "9A" }
  @{ en = "Ganne Yehoshua"; he = "גני יהושע"; section = "10A" }
  @{ en = "Hadar Yosef"; he = "הדר יוסף"; section = "10A" }
  @{ en = "Pinhas Rosen Bridge"; he = "גשר פנחס רוזן"; section = "11A" }
  @{ en = "HaBarzel"; he = "הברזל"; section = "11A" }
  @{ en = "HaNehoshet"; he = "הנחושת"; section = "11A" }
  @{ en = "Dvora HaNevi'a"; he = "דבורה הנביאה"; section = "11A" }
  @{ en = "Neve Sharet"; he = "נווה שרת"; section = "11A" }
  @{ en = "Abba Eban"; he = "אבא אבן"; section = "12A" }
  @{ en = "HaHoshlim"; he = "החושלים"; section = "12A" }
  @{ en = "Altneuland"; he = "אלטנוילנד"; section = "12A" }
  @{ en = "Hof HaTkhelet North"; he = "חוף התכלת צפון"; section = "12A" }
  @{ en = "Hof HaTkhelet South"; he = "חוף התכלת דרום"; section = "12A" }
  @{ en = "Unichman"; he = "יוניצ'מן"; section = "12A" }
  @{ en = "Miryam Yalan-Shteklis"; he = "מרים ילן שטקליס"; section = "12A" }
  @{ en = "Miryam Ben-Porat"; he = "מרים בן פורת"; section = "12A" }
  @{ en = "Shoshanna Persitz"; he = "שושנה פרסיץ"; section = "12A" }
  @{ en = "Azorey Hen"; he = "אזורי חן"; section = "12A" }
  @{ en = "Propes"; he = "פרופס"; section = "12A" }
  @{ en = "Nofe Yam"; he = "נופי ים"; section = "12A" }
  @{ en = "Einstein"; he = "איינשטיין"; section = "12A" }
  @{ en = "Levi Eshkol"; he = "לוי אשכול"; section = "12A" }
  @{ en = "Zohara Leviatov"; he = "זהרה לביטוב"; section = "12A" }
)

function Normalize-StopName($s) {
  if (-not ($s -is [string])) { return "" }
  return ($s.ToLower() -replace '[^a-z0-9]', '')
}

function Get-StopsData($ws) {
  $used = $ws.UsedRange
  $maxRow = $used.Row + $used.Rows.Count - 1
  $maxCol = $used.Column + $used.Columns.Count - 1

  $cols = @{}
  foreach ($key in $STOPS_GROUP_LABELS.Keys) {
    $label = $STOPS_GROUP_LABELS[$key]
    for ($c = 1; $c -le $maxCol; $c++) {
      $r3 = $ws.Cells.Item(3, $c).Value2
      if ($r3 -is [string] -and $r3.Trim() -eq $label) { $cols[$key] = $c; break }
    }
    if (-not $cols.ContainsKey($key)) {
      Write-Output "WARN Stops sheet: column not found for '$label'"
    }
  }
  if (-not $cols.ContainsKey("location") -or -not $cols.ContainsKey("name")) { return @() }

  function Get-StopPct($ws, $r, $cols, $key) {
    if (-not $cols.ContainsKey($key)) { return 0 }
    $v = $ws.Cells.Item($r, $cols[$key]).Value2
    if ($v -is [double]) { return [math]::Round(100.0 * $v, 1) }
    return 0
  }

  # Same as Get-StopPct but returns $null (not 0) when there's no numeric value -
  # used for "CW Station - Platform Surfacing" so the UI can show "אין נתון"
  # instead of implying a genuine 0% for stations the report doesn't cover.
  function Get-StopPctOrNull($ws, $r, $cols, $key) {
    if (-not $cols.ContainsKey($key)) { return $null }
    $v = $ws.Cells.Item($r, $cols[$key]).Value2
    if ($v -is [double]) { return [math]::Round(100.0 * $v, 1) }
    return $null
  }

  $rawByNorm = @{}
  for ($r = 5; $r -le $maxRow; $r++) {
    $name = $ws.Cells.Item($r, $cols["name"]).Value2
    if (-not ($name -is [string]) -or [string]::IsNullOrWhiteSpace($name)) { continue }
    $loc = $ws.Cells.Item($r, $cols["location"]).Value2
    $loc = if ($loc -is [string]) { $loc.Trim() } else { "" }
    if ($loc -notmatch '^\d{1,2}A$') { continue }

    $secondary = Get-StopPct $ws $r $cols "secondary"
    $cslm      = Get-StopPct $ws $r $cols "cslm"
    $sig       = Get-StopPct $ws $r $cols "sig"
    $com       = Get-StopPct $ws $r $cols "com"
    $surfacing = Get-StopPctOrNull $ws $r $cols "surfacing"

    # Flat, equally-weighted average across the 4 discipline percentages (each
    # weighted 1/4) - per Omri's request, replacing Excel's own "Stops -
    # General" column (a differently-weighted formula) as this station's
    # headline %, so it matches the average of the 4 rows shown in the detail
    # panel. "surfacing" (added 2026-07-19) is shown as its own extra row in
    # the detail panel and deliberately left out of this average.
    $entry = @{
      overall   = [math]::Round((($secondary + $cslm + $sig + $com) / 4.0), 1)
      secondary = $secondary
      cslm      = $cslm
      sig       = $sig
      com       = $com
      surfacing = $surfacing
    }
    $rawByNorm[(Normalize-StopName $name.Trim())] = $entry
  }

  $out = @()
  foreach ($m in $STOPS_MASTER_LIST) {
    $norm = Normalize-StopName $m.en
    $raw = $rawByNorm[$norm]
    if ($raw) {
      $pct = $raw.overall
      $details = [ordered]@{ secondary = $raw.secondary; cslm = $raw.cslm; sig = $raw.sig; com = $raw.com; surfacing = $raw.surfacing }
    } else {
      $pct = 0
      $details = [ordered]@{ secondary = 0; cslm = 0; sig = 0; com = 0; surfacing = $null }
    }
    $out += [ordered]@{ en = $m.en; he = $m.he; section = $m.section; pct = $pct; details = $details }
  }
  return $out
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

function Get-JunctionDisplayName($rawName) {
  # Column D packs a human-readable description followed by one or two
  # internal engineering codes, e.g. " Ibn Gabirol - Avenue 16 (HR2200) 0
  # Phases CR-G5-30 G5-6" or " Sports Center Parking CR-G4-15". Per Omri's
  # request (2026-07-16), only the descriptive part should show - strip
  # everything from the first "CR-G..." code onward, then trim any
  # leftover trailing separator (space/comma/dash) the code left behind.
  $stripped = $rawName -replace '\s*CR-G.*$', ''
  return ($stripped -replace '[\s,\-]+$', '').Trim()
}

function Add-JunctionData($ws, $lotJunctions) {
  # "Junctions" sheet: one row per junction. Column B ("A section") is the
  # 1A-12A section code, column D ("Name") is the junction's English name
  # (descriptive part kept, internal CR-G.. code stripped - see
  # Get-JunctionDisplayName), column G ("Execution %") is that junction's
  # own completion fraction (0-1). A trailing "Average" summary row and
  # blank rows have an empty column B and are skipped naturally.
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
    $name = if ($name -is [string]) { Get-JunctionDisplayName $name.Trim() } else { "" }

    if (-not $lotJunctions.ContainsKey($loc)) {
      $lotJunctions[$loc] = @{ total = 0; completed = 0; inProgress = 0; completedNames = @(); inProgressNames = @() }
    }
    $lotJunctions[$loc].total++
    if ($pct -ge 1) {
      $lotJunctions[$loc].completed++
      if ($name) { $lotJunctions[$loc].completedNames += [ordered]@{ name = $name; section = $loc } }
    }
    elseif ($pct -gt 0) {
      $lotJunctions[$loc].inProgress++
      if ($name) { $lotJunctions[$loc].inProgressNames += [ordered]@{ name = $name; section = $loc } }
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
$stopsResults = @{}

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
    foreach ($t in $EXTRA_DISCIPLINES) { $mainLookup[$t] = $t }
    $discBlocks = Find-DisciplineBlocks $ws $maxCol $mainLookup
    foreach ($target in ($TARGET_DISCIPLINES + $EXTRA_DISCIPLINES)) {
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
        foreach ($target in ($TARGET_DISCIPLINES + $EXTRA_DISCIPLINES)) {
          if (-not $lotDiscAgg[$loc].ContainsKey($target)) { continue }
          $entry = $lotDiscAgg[$loc][$target]
          if ($entry.sumPlanned -le 0) { continue }
          $disciplines[$target] = [ordered]@{
            pct     = [math]::Round(100.0 * $entry.sumActual / $entry.sumPlanned, 1)
            actual  = [math]::Round($entry.sumActual, 1)
            planned = [math]::Round($entry.sumPlanned, 1)
            unit    = $(if ($entry.unit) { $entry.unit } else { "" })
          }
          # $EXTRA_DISCIPLINES (CW Track surfacing) is shown as its own detail-panel
          # row but deliberately excluded from the section's headline % - see the
          # $EXTRA_DISCIPLINES comment above.
          if ($TARGET_DISCIPLINES -contains $target) {
            $sumActualAll += $entry.sumActual
            $sumPlannedAll += $entry.sumPlanned
          }
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

    # Stops: one row per station, matched against the canonical 62-station
    # Hebrew/section list (see $STOPS_MASTER_LIST above).
    $stopsSheet = $null
    foreach ($s in $wb.Worksheets) { if ($s.Name -eq "Stops") { $stopsSheet = $s.Name } }
    $stopsList = @()
    if ($stopsSheet) { $stopsList = Get-StopsData $wb.Worksheets.Item($stopsSheet) }
    $stopsResults[$dateStr] = $stopsList

    Write-Output "OK $dateStr <- $($f.Name)  lots=$($lotFinal.Count) disciplines=$($discBlocks.Count)/$($TARGET_DISCIPLINES.Count) tunnel=$($tunnelBlocks.Count -as [string]) junctionSections=$($lotJunctions.Count) depots=$($depotsForDate.Count) ttr=$($ttrList.Count) stops=$($stopsList.Count)"

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

$stopsSortedDates = $stopsResults.Keys | Sort-Object
$stopsObjAll = [ordered]@{}
foreach ($d in $stopsSortedDates) { $stopsObjAll[$d] = $stopsResults[$d] }
$stopsJson = $stopsObjAll | ConvertTo-Json -Depth 6

$js = "// Auto-generated by update_progress_data.ps1 - do not edit by hand`r`nconst PROGRESS_DATA = $json;`r`nconst DEPOT_DATA = $depotJson;`r`nconst TTR_DATA = $ttrJson;`r`nconst STOPS_DATA = $stopsJson;`r`n"
Set-Content -Path $OutJs -Value $js -Encoding utf8
Write-Output "WROTE: $OutJs"
