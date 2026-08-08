$ErrorActionPreference = 'Stop'
$out = "C:\Users\harry\aston-martin-financing-analysis\AML_Liquidity_Model.xlsx"
if (Test-Path $out) { Remove-Item $out -Force }

$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false
$xl.DisplayAlerts = $false
$xl.SheetsInNewWorkbook = 1
$wb = $xl.Workbooks.Add()

$BLUE  = 15773696
$AMBER = 49407
$RED   = 255
$HDRBG = 6299648

$NUM  = '#,##0.0;(#,##0.0)'
$NUM0 = '#,##0;(#,##0)'
$PCT  = '0.0%'
$XX   = '0.00"x"'

function SetC($ws, $r, $c, $v, [switch]$Bold, $NumFmt, $FontColor, [switch]$Italic) {
  $cell = $ws.Cells.Item($r, $c)
  if ($v -ne $null) {
    if ($v -is [string]) {
      if ($v.StartsWith('=')) { $cell.Formula = [string]$v } else { $cell.Value2 = [string]$v }
    } else {
      $cell.Value2 = [double]$v
    }
  }
  if ($Bold) { $cell.Font.Bold = $true }
  if ($Italic) { $cell.Font.Italic = $true }
  if ($NumFmt) { $cell.NumberFormat = $NumFmt }
  if ($FontColor -ne $null) { $cell.Font.Color = $FontColor }
}
function SectionHdr($ws, $r, $text, $lastCol) {
  $rng = $ws.Range($ws.Cells.Item($r,1), $ws.Cells.Item($r,$lastCol))
  $rng.Interior.Color = $HDRBG
  $rng.Font.Color = 16777215
  $rng.Font.Bold = $true
  $ws.Cells.Item($r,1).Value2 = $text
}

$yrs = @("FY2025A","FY2026E","FY2027E","FY2028E","FY2029E","FY2030E","FY2031E")

# ===================== README =====================
$ws = $wb.Worksheets.Item(1)
$ws.Name = "README"
SetC $ws 1 1 "ASTON MARTIN LAGONDA - LIQUIDITY AND CAPITAL STRUCTURE MODEL" -Bold
$ws.Cells.Item(1,1).Font.Size = 14
$readme = @(
 @(3,  "Purpose",       "Credit / capital-structure analysis of the July 2026 GBP 550m private-credit financing. NOT a valuation model."),
 @(4,  "Question",      "Did the financing solve the liquidity problem, or postpone further deleveraging / an equity raise?"),
 @(6,  "PHASE",         "Phase 2 (model build) complete. Phase 3 (downside scenario) NOT built - parameters await sign-off."),
 @(8,  "SOURCING RULE", "Every hard-coded figure traces to an AML primary document. See SOURCES.md."),
 @(9,  "",              "Press reporting (bondholder friction / drop-down) is NOT an input to this model anywhere."),
 @(11, "COLOUR CODE",   ""),
 @(12, "  BLUE",        "Sourced from an AML primary document."),
 @(13, "  AMBER",       "PROVISIONAL judgement - awaiting sign-off. Drives the conclusion. Edit on Assumptions tab."),
 @(14, "  RED",         "Unsourced / disclosed limitation."),
 @(15, "  BLACK",       "Calculated - do not overwrite."),
 @(17, "TABS",          ""),
 @(18, "  Assumptions", "All inputs. The ONLY tab you should edit."),
 @(19, "  Debt",        "Pre vs post capital structure, interest schedule, reconciliation to reported net debt."),
 @(20, "  Model",       "P and L, free cash flow, net debt and liquidity roll-forward FY2025A to FY2031E."),
 @(21, "  Outputs",     "Credit metrics dashboard and the headline answer."),
 @(22, "  Maturity",    "Debt maturity profile."),
 @(24, "KEY FLAGS",     ""),
 @(25, "  A1 Txn costs","Not disclosed by AML. Band 30/40/50; central GBP 40m reproduces the disclosed c.GBP 340m."),
 @(26, "  A2 DDTL",     "GBP 100m delayed draw EXCLUDED from pro forma liquidity (CFO confirmed). Memo item only."),
 @(27, "  Covenant",    "SSTL minimum-liquidity covenant tested monthly from Aug-2026 but the LEVEL is not disclosed."),
 @(28, "  FY27-31",     "AML gives NO quantified guidance beyond FY2026. All FY2027+ drivers are provisional."),
 @(29, "  REFI RATE",   "20.16pct is a market-implied DISTRESS yield, NOT a forecast new-issue coupon. This caveat must travel"),
 @(30, "",              "with the number into the memo and deck - it is not sufficient for it to sit on the YTM tab alone."),
 @(32, "  ATTRIBUTION", "See the Attribution tab: the leverage reversal is driven MAINLY by the margin cap, not the refi rate.")
)
foreach ($r in $readme) { SetC $ws $r[0] 1 $r[1] -Bold; SetC $ws $r[0] 2 $r[2] }
$ws.Cells.Item(12,1).Font.Color = $BLUE
$ws.Cells.Item(13,1).Font.Color = $AMBER
$ws.Cells.Item(14,1).Font.Color = $RED
$ws.Columns.Item(1).ColumnWidth = 20
$ws.Columns.Item(2).ColumnWidth = 108

# ===================== ASSUMPTIONS =====================
$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
$ws.Name = "Assumptions"
SetC $ws 1 1 "ASSUMPTIONS - the only tab to edit" -Bold
$ws.Cells.Item(1,1).Font.Size = 13
SetC $ws 3 1 "SCENARIO SWITCH" -Bold
SetC $ws 3 2 1 -NumFmt "0" -FontColor $AMBER
SetC $ws 3 3 '=IF(B3=1,"BASE CASE","DOWNSIDE CASE")' -Bold
SetC $ws 4 1 "1 = Base. 2 = Downside (single combined stress: volumes, gross margin and SONIA together from FY2027)." -Italic
SetC $ws 6 1 "LEGEND" -Bold
SetC $ws 7 1 "BLUE = sourced from AML primary document" -FontColor $BLUE
SetC $ws 8 1 "AMBER = PROVISIONAL, awaiting sign-off" -FontColor $AMBER
SetC $ws 9 1 "RED = unsourced / limitation" -FontColor $RED

SectionHdr $ws 11 "ACTIVE DRIVERS - CALCULATED, switch with B3. Edit the BASE and DOWNSIDE blocks below." 9
SetC $ws 12 1 "Driver" -Bold
for ($i=0; $i -lt 7; $i++) { SetC $ws 12 (2+$i) $yrs[$i] -Bold }
SetC $ws 12 9 "Basis" -Bold

$actLab = @("Total wholesale volumes","Blended ASP (GBP k)","Gross margin pct","Adjusted opex ex D and A","Adjusted D and A","Capex","Working capital movement","Lease payments","Cash tax","Other cash items")
for ($k=0; $k -lt 10; $k++) {
  $ar = 13 + $k; $br = 42 + $k; $dr = 57 + $k
  SetC $ws $ar 1 $actLab[$k]
  for ($i=0; $i -lt 7; $i++) {
    $c = [char](66+$i)
    if ($ar -eq 15) { $fmt = $PCT } elseif ($ar -eq 13) { $fmt = $NUM0 } else { $fmt = $NUM }
    SetC $ws $ar (2+$i) ('=IF($B$3=1,{0}{1},{0}{2})' -f $c,$br,$dr) -NumFmt $fmt
  }
  SetC $ws $ar 9 ("CALCULATED - base row {0} / downside row {1}" -f $br,$dr) -Italic
}

$drv = @(
 @(13,"Total wholesale volumes",      5448,   5900,  6100,  6300,  6400,  6500,  6600,  "FY25A sourced [AR25]. FY26+ PROVISIONAL - no volume guidance is given.", "A"),
 @(14,"Blended ASP (GBP k)",          230.9,  271.0, 274.0, 277.0, 280.0, 283.0, 286.0, "FY25A = revenue/volume. FY26+ PROVISIONAL (c.500 Valhalla in FY26 is sourced).", "A"),
 @(15,"Gross margin pct",             0.294,  0.350, 0.355, 0.355, 0.355, 0.355, 0.355, "FY25A sourced. H1-26A 33.8pct sourced. FY26+ PROVISIONAL, set so EBITDA margin caps at c.17.1pct.", "A"),
 @(16,"Adjusted opex ex D and A",     262,    295,   305,   320,   330,   340,   350,   "FY25A sourced. FY26 below 300 = GUIDANCE. FY27+ PROVISIONAL, calibrated to the 17.1pct margin cap.", "A"),
 @(17,"Adjusted D and A",             297,    387.5, 390,   380,   370,   360,   350,   "FY25A sourced. FY26 375-400 = GUIDANCE (midpoint). FY27+ PROVISIONAL.", "A"),
 @(18,"Capex",                        341,    300,   350,   350,   350,   350,   300,   "FY25A sourced. FY26 c.300 = GUIDANCE. FY26-30 sums to c.GBP 1.7bn = GUIDANCE.", "B"),
 @(19,"Working capital movement",     0,      -20,   -10,   -10,   -10,   -10,   -10,   "PROVISIONAL.", "A"),
 @(20,"Lease payments",               0,      -10,   -10,   -10,   -10,   -10,   -10,   "FY25A lease payments GBP 10.0m sourced. Held flat - PROVISIONAL.", "A"),
 @(21,"Cash tax",                     0,      -5,    -5,    -5,    -5,    -5,    -5,    "PROVISIONAL (loss-making; DTA written to nil at FY25).", "A"),
 @(22,"Other cash items",             0,      0,     0,     0,     0,     0,     0,     "PROVISIONAL.", "A")
)
SectionHdr $ws 40 "BASE CASE INPUTS - edit here" 9
SetC $ws 41 1 "Driver" -Bold
for ($i=0; $i -lt 7; $i++) { SetC $ws 41 (2+$i) $yrs[$i] -Bold }
SetC $ws 41 9 "Basis" -Bold
foreach ($d in $drv) {
  $row = $d[0] + 29    # 13->42 ... 22->51
  SetC $ws $row 1 $d[1]
  for ($i=0; $i -lt 7; $i++) {
    if ($row -eq 44) { $fmt = $PCT } elseif ($row -eq 42) { $fmt = $NUM0 } else { $fmt = $NUM }
    if ($d[10] -eq "B") { $col = $BLUE } elseif ($i -eq 0) { $col = $BLUE } else { $col = $AMBER }
    SetC $ws $row (2+$i) $d[2+$i] -NumFmt $fmt -FontColor $col
  }
  SetC $ws $row 9 $d[9] -Italic
}
SetC $ws 52 1 "SONIA"
for ($i=0; $i -lt 7; $i++) { SetC $ws 52 (2+$i) 0.0372 -NumFmt '0.00%' -FontColor $BLUE }
SetC $ws 52 9 "BoE Bank Rate 3.75pct (Jul-2026) less c.3bp. Forward curve not retrievable (403)." -Italic

SectionHdr $ws 55 "DOWNSIDE CASE INPUTS - single combined stress, anchored on management's reverse stress test" 9
SetC $ws 56 1 "Driver" -Bold
for ($i=0; $i -lt 7; $i++) { SetC $ws 56 (2+$i) $yrs[$i] -Bold }
SetC $ws 56 9 "Basis" -Bold
$down = @(
 @(57,"Total wholesale volumes",   5448,  5900,  4697,  4851,  4928,  5005,  5082,  "STRESSED FY27+: -23pct vs base = management's -25pct CORE stress with Specials held flat (Specials c.8.5pct of units).", "S"),
 @(58,"Blended ASP (GBP k)",       230.9, 271.0, 289.07,292.24,295.40,298.57,301.73,"STRESSED FY27+: +5.5pct vs base. Mix effect ONLY - Specials become a larger share of a smaller book. Net revenue c.-18.8pct.", "S"),
 @(59,"Gross margin pct",          0.294, 0.350, 0.320, 0.320, 0.320, 0.320, 0.320, "STRESSED FY27+: 32.0pct, between the FY2025 trough (29.4pct) and H1-26 (33.8pct), both sourced.", "S"),
 @(60,"Adjusted opex ex D and A",  262,   295,   305,   320,   330,   340,   350,   "UNCHANGED vs base - fixed cost base. This is the operating-deleverage effect.", "U"),
 @(61,"Adjusted D and A",          297,   387.5, 390,   380,   370,   360,   350,   "UNCHANGED vs base.", "U"),
 @(62,"Capex",                     341,   300,   350,   350,   350,   350,   300,   "UNCHANGED vs base - management's reverse stress test EXCLUDES mitigating actions.", "U"),
 @(63,"Working capital movement",  0,     -20,   -10,   -10,   -10,   -10,   -10,   "UNCHANGED vs base - no additional judgement layered on.", "U"),
 @(64,"Lease payments",            0,     -10,   -10,   -10,   -10,   -10,   -10,   "UNCHANGED vs base.", "U"),
 @(65,"Cash tax",                  0,     -5,    -5,    -5,    -5,    -5,    -5,    "UNCHANGED vs base.", "U"),
 @(66,"Other cash items",          0,     0,     0,     0,     0,     0,     0,     "UNCHANGED vs base.", "U"),
 @(67,"SONIA",                     0.0372,0.0372,0.0525,0.0525,0.0525,0.0525,0.0525,"STRESSED FY27+: 5.25pct (+153bps). JUDGEMENT - no forward curve obtainable. Low materiality: c.GBP 4.5m per 100bps.", "S")
)
foreach ($d in $down) {
  $row = $d[0]
  SetC $ws $row 1 $d[1]
  for ($i=0; $i -lt 7; $i++) {
    if ($row -eq 59) { $fmt = $PCT } elseif ($row -eq 67) { $fmt = '0.00%' } elseif ($row -eq 57) { $fmt = $NUM0 } else { $fmt = $NUM }
    if ($d[10] -eq "S") { $col = $RED } else { $col = $AMBER }
    SetC $ws $row (2+$i) $d[2+$i] -NumFmt $fmt -FontColor $col
  }
  SetC $ws $row 9 $d[9] -Italic
}
SetC $ws 69 1 "THREE STRESS LEVERS ONLY (red): volumes, gross margin, SONIA - applied together from FY2027, per the brief." -Bold
SetC $ws 70 1 "FY2025A and FY2026E are NOT stressed: FY25 is actual and FY26 is covered by company guidance."
SetC $ws 71 1 "NO-RESPONSE CASE: capex and opex are deliberately held flat because management's disclosed reverse stress test" -Bold
SetC $ws 72 1 "excludes mitigating actions. Real-world mitigation (capex deferral, cost action) would soften this materially -" -Bold
SetC $ws 73 1 "note that qualitatively in the memo, but do NOT model it." -Bold

SectionHdr $ws 24 "RATES AND FX" 9
SetC $ws 25 1 "Driver" -Bold
for ($i=0; $i -lt 7; $i++) { SetC $ws 25 (2+$i) $yrs[$i] -Bold }
SetC $ws 25 9 "Basis" -Bold
SetC $ws 26 1 "SONIA"
for ($i=0; $i -lt 7; $i++) { $c=[char](66+$i); SetC $ws 26 (2+$i) ('=IF($B$3=1,{0}52,{0}67)' -f $c) -NumFmt '0.00%' }
SetC $ws 26 9 "CALCULATED - base row 52 / downside row 67." -Italic
$rates = @(
 @(27,"USD per GBP",           1.3253,1.3253,1.3253,1.3253,1.3253,1.3253,1.3253,"SOURCED (derived): USD 1,050.0m / GBP 792.3m at 30-Jun-26 [H1].","B"),
 @(28,"SSTL margin over SONIA",0.0675,0.0675,0.0675,0.0675,0.0675,0.0675,0.0675,"SOURCED: SONIA + 6.75pct [H1].","B"),
 @(29,"Refi rate on 2029 notes",0.2016,0.2016,0.2016,0.2016,0.2016,0.2016,0.2016,"COMPUTED blended YTM from the 30-Jun-26 market price (79.57pct of par). CAVEAT: this is a market-implied DISTRESS yield embedding default probability, NOT a forecast new-issue coupon - never quote it without that caveat. See YTM tab.","B")
)
foreach ($d in $rates) {
  $row = $d[0]
  SetC $ws $row 1 $d[1]
  for ($i=0; $i -lt 7; $i++) {
    if ($row -eq 27) { $fmt = '0.0000' } else { $fmt = '0.00%' }
    if ($d[10] -eq "B") { $col = $BLUE } else { $col = $AMBER }
    SetC $ws $row (2+$i) $d[2+$i] -NumFmt $fmt -FontColor $col
  }
  SetC $ws $row 9 $d[9] -Italic
}

SectionHdr $ws 31 "SWITCHES AND SINGLE INPUTS" 9
SetC $ws 32 1 "Transaction cost case (1=Low 2=Central 3=High)"
SetC $ws 32 2 2 -NumFmt "0" -FontColor $AMBER
SetC $ws 32 9 "User-approved band 30/40/50. Central reproduces the disclosed c.GBP 340m." -Italic
SetC $ws 33 1 "Transaction costs (GBP m)"
SetC $ws 33 2 '=CHOOSE(B32,30,40,50)' -NumFmt $NUM -Bold
SetC $ws 33 9 "A1 - NOT DISCLOSED by AML. Unverified pending FY2026 Annual Report." -Italic -FontColor $RED
SetC $ws 34 1 "DDTL drawn (0=No 1=Yes)"
SetC $ws 34 2 '=IF($B$3=1,0,1)' -NumFmt "0"
SetC $ws 34 9 "CALCULATED - undrawn in base, DRAWN in downside (a company under liquidity stress draws committed capacity). A2: excluded from the disclosed c.GBP 340m (CFO confirmed)." -Italic
SetC $ws 35 1 "Refinance 2029 notes at maturity (0/1)"
SetC $ws 35 2 1 -NumFmt "0" -FontColor $AMBER
SetC $ws 35 9 "Set to 0 to expose the raw 2029 funding wall." -Italic
SetC $ws 36 1 "Minimum liquidity covenant (GBP m)"
SetC $ws 36 2 100 -NumFmt $NUM -FontColor $RED
SetC $ws 36 9 "QUANTUM NOT DISCLOSED. Covenant exists, tested monthly from Aug-26 [H1]." -Italic -FontColor $RED
SetC $ws 37 1 "Inventory financing assumed rolled (0/1)"
SetC $ws 37 2 1 -NumFmt "0" -FontColor $AMBER
SetC $ws 37 9 "GBP 39.1m repayable Oct-26 (GBP 40.0m gross) [H1]. Assumed rolled." -Italic

$ws.Columns.Item(1).ColumnWidth = 38
for ($c=2; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 11 }
$ws.Columns.Item(9).ColumnWidth = 76

# ===================== DEBT =====================
$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
$ws.Name = "Debt"
SetC $ws 1 1 "CAPITAL STRUCTURE - PRE vs POST TRANSACTION (GBP m, nominal principal)" -Bold
$ws.Cells.Item(1,1).Font.Size = 13
SetC $ws 3 1 "Instrument" -Bold
SetC $ws 3 2 "Pre 30-Jun-26" -Bold
SetC $ws 3 3 "Transaction" -Bold
SetC $ws 3 4 "Post pro forma" -Bold
SetC $ws 3 5 "Coupon" -Bold
SetC $ws 3 6 "Maturity" -Bold
SetC $ws 3 7 "Source" -Bold
$stack = @(
 @(4, "USD SSN (USD 1,050.0m at 10.0pct)", 792.3, '0',                      "10.0pct",     "Mar-2029",  "[H1] fair value note"),
 @(5, "GBP SSN (GBP 465.0m at 10.375pct)", 465.0, '0',                      "10.375pct",   "Mar-2029",  "[H1] fair value note"),
 @(6, "GBP SSN (GBP 100.0m at 10.375pct)", 100.0, '0',                      "10.375pct",   "Mar-2029",  "[H1] fair value note"),
 @(7, "SSTL (GBP 450m)",                   0,     '450',                    "SONIA+6.75",  "Jul-2031",  "[H1] going concern"),
 @(8, "DDTL (GBP 100m delayed draw)",      0,     '=Assumptions!$B$34*100', "SONIA+6.75",  "Jul-2031",  "[H1] memo, excl. liquidity"),
 @(9, "Super senior RCF (drawn)",          163.0, '-163',                   "SONIA-based", "cancelled", "[H1] GBP 163.0m drawn"),
 @(10,"YTC committed facility (drawn)",    20.0,  '-20',                    "n/d",         "cancelled", "[H1] GBP 20m of GBP 50m"),
 @(11,"Other bank loans and overdrafts",   6.5,   '0',                      "n/d",         "various",   "[H1] balance to GBP 168.8m"),
 @(12,"Inventory financing",               39.1,  '0',                      "n/d",         "Oct-2026",  "[H1]"),
 @(13,"Lease liabilities (IFRS 16)",       89.7,  '0',                      "n/a",         "various",   "[H1]")
)
foreach ($s in $stack) {
  $row = $s[0]
  SetC $ws $row 1 $s[1]
  SetC $ws $row 2 $s[2] -NumFmt $NUM -FontColor $BLUE
  SetC $ws $row 3 $s[3] -NumFmt $NUM
  SetC $ws $row 4 ('=B{0}+C{0}' -f $row) -NumFmt $NUM
  SetC $ws $row 5 $s[4]
  SetC $ws $row 6 $s[5]
  SetC $ws $row 7 $s[6] -Italic
}
SetC $ws 14 1 "Gross debt (nominal)" -Bold
SetC $ws 14 2 '=SUM(B4:B13)' -NumFmt $NUM -Bold
SetC $ws 14 3 '=SUM(C4:C13)' -NumFmt $NUM -Bold
SetC $ws 14 4 '=SUM(D4:D13)' -NumFmt $NUM -Bold
SetC $ws 15 1 "Unamortised transaction fees"
SetC $ws 15 2 -14.5 -NumFmt $NUM -FontColor $BLUE
SetC $ws 15 7 "[H1] notes 12.1 + RCF 0.7 + YTC 1.7" -Italic
SetC $ws 16 1 "Gross debt (book)" -Bold
SetC $ws 16 2 '=B14+B15' -NumFmt $NUM -Bold
SetC $ws 17 1 "CHECK - reported gross debt [H1]"
SetC $ws 17 2 1661.1 -NumFmt $NUM -FontColor $BLUE
SetC $ws 18 1 "Difference"
SetC $ws 18 2 '=B16-B17' -NumFmt $NUM
SetC $ws 18 3 '=IF(ABS(B18)<0.05,"TIES","CHECK")' -Bold
SetC $ws 20 1 "Cash" -Bold
SetC $ws 20 2 114.9 -NumFmt $NUM -FontColor $BLUE
SetC $ws 20 3 '=450+Assumptions!$B$34*100-163-20-Assumptions!$B$33' -NumFmt $NUM
SetC $ws 20 4 '=B20+C20' -NumFmt $NUM -Bold
SetC $ws 21 1 "Cash not available for short-term use"
SetC $ws 21 2 1.5 -NumFmt $NUM -FontColor $BLUE
SetC $ws 21 4 '=B21' -NumFmt $NUM
SetC $ws 22 1 "Net debt (nominal basis)" -Bold
SetC $ws 22 2 '=B14-B20-B21' -NumFmt $NUM -Bold
SetC $ws 22 4 '=D14-D20-D21' -NumFmt $NUM -Bold
SetC $ws 23 1 "Available facilities"
SetC $ws 23 2 30.3 -NumFmt $NUM -FontColor $BLUE
SetC $ws 23 4 0 -NumFmt $NUM
SetC $ws 23 7 "RCF and YTC both cancelled [H1]" -Italic
SetC $ws 24 1 "TOTAL LIQUIDITY" -Bold
SetC $ws 24 2 '=B20+B23' -NumFmt $NUM -Bold
SetC $ws 24 4 '=D20+D23' -NumFmt $NUM -Bold
SetC $ws 25 1 "CHECK vs disclosed"
SetC $ws 25 2 145.2 -NumFmt $NUM -FontColor $BLUE
SetC $ws 25 4 340 -NumFmt $NUM -FontColor $BLUE
SetC $ws 26 1 "Difference"
SetC $ws 26 2 '=B24-B25' -NumFmt $NUM
SetC $ws 26 4 '=D24-D25' -NumFmt $NUM
SetC $ws 27 1 "MEMO - undrawn DDTL capacity (NOT in liquidity)" -Bold
SetC $ws 27 4 '=(1-Assumptions!$B$34)*100' -NumFmt $NUM -FontColor $AMBER
SetC $ws 28 1 "NOTE: the 'CHECK vs disclosed' above is calibrated to the BASE case (DDTL undrawn), which is how AML disclosed it." -Italic
SetC $ws 29 1 "In the DOWNSIDE the DDTL is drawn, so post-transaction cash is c.GBP 100m higher by construction - the check will not tie there." -Italic

SectionHdr $ws 30 "CASH INTEREST SCHEDULE (GBP m)" 8
SetC $ws 31 1 "Line" -Bold
for ($i=1; $i -lt 7; $i++) { SetC $ws 31 (1+$i) $yrs[$i] -Bold }
SetC $ws 32 1 "2029 notes outstanding - year fraction"
$nf = @(1,1,1,0.25,0,0)
for ($i=0; $i -lt 6; $i++) { SetC $ws 32 (2+$i) $nf[$i] -NumFmt "0.00" -FontColor $BLUE }
SetC $ws 32 8 "Mature Mar-2029 [H1], so 3/12 of FY2029" -Italic
SetC $ws 33 1 "SSTL - year fraction"
$sf = @(0.44,1,1,1,1,1)
for ($i=0; $i -lt 6; $i++) { SetC $ws 33 (2+$i) $sf[$i] -NumFmt "0.00" -FontColor $BLUE }
SetC $ws 33 8 "Drawn 22-Jul-2026, so c.5.3 months of FY2026" -Italic
SetC $ws 34 1 "USD SSN interest"
SetC $ws 35 1 "GBP SSN interest"
SetC $ws 36 1 "SSTL interest"
SetC $ws 37 1 "DDTL interest"
SetC $ws 38 1 "Refinancing debt interest (post Mar-29)"
SetC $ws 39 1 "Other interest (RCF/YTC/inventory/leases)"
SetC $ws 40 1 "Interest received"
SetC $ws 41 1 "Subtotal - accrual basis" -Bold
SetC $ws 42 1 "FY2026 calibration to guided cash interest"
SetC $ws 43 1 "NET CASH INTEREST" -Bold
for ($i=0; $i -lt 6; $i++) {
  $d = [char](66+$i)   # Debt col B..G  = FY2026..FY2031
  $a = [char](67+$i)   # Assumptions col C..H = FY2026..FY2031
  SetC $ws 34 (2+$i) ('=1050*0.10/Assumptions!{1}27*{0}32' -f $d,$a) -NumFmt $NUM
  SetC $ws 35 (2+$i) ('=565*0.10375*{0}32' -f $d) -NumFmt $NUM
  SetC $ws 36 (2+$i) ('=450*(Assumptions!{1}26+Assumptions!{1}28)*{0}33' -f $d,$a) -NumFmt $NUM
  SetC $ws 37 (2+$i) ('=Assumptions!$B$34*100*(Assumptions!{1}26+Assumptions!{1}28)*{0}33' -f $d,$a) -NumFmt $NUM
  SetC $ws 38 (2+$i) ('=Assumptions!$B$35*1357.3*Assumptions!{1}29*(1-{0}32)' -f $d,$a) -NumFmt $NUM
  SetC $ws 40 (2+$i) 2 -NumFmt $NUM
  SetC $ws 41 (2+$i) ('=SUM({0}34:{0}39)-{0}40' -f $d) -NumFmt $NUM -Bold
  if ($i -gt 0) { SetC $ws 42 (2+$i) 0 -NumFmt $NUM }
  SetC $ws 43 (2+$i) ('={0}41+{0}42' -f $d) -NumFmt $NUM -Bold
}
$oth = @(15,8,8,8,8,8)
for ($i=0; $i -lt 6; $i++) { SetC $ws 39 (2+$i) $oth[$i] -NumFmt $NUM -FontColor $AMBER }
SetC $ws 39 8 "FY26 includes part-year RCF/YTC. PROVISIONAL." -Italic
SetC $ws 38 8 "At 20.16pct = market-implied DISTRESS yield, NOT a forecast new-issue coupon. Never quote without this caveat. See YTM tab." -Italic -FontColor $RED
SetC $ws 42 2 '=B44-B41' -NumFmt $NUM -FontColor $AMBER
SetC $ws 42 8 "FY2026 ONLY - calibrates to company guidance. NOT a silent override." -Italic
SetC $ws 44 1 "GUIDANCE - FY2026 net cash interest c.GBP 160m [H1]"
SetC $ws 44 2 160 -NumFmt $NUM -FontColor $BLUE
SetC $ws 45 1 "Model vs guidance (should be nil after calibration)"
SetC $ws 45 2 '=B43-B44' -NumFmt $NUM
SetC $ws 47 1 "ACCRUAL vs CASH: rows 34-41 accrue a full year of coupons. FY2026 is calibrated down to the guided c.GBP 160m cash figure," -Italic
SetC $ws 48 1 "because guidance reflects actual payment dates. FY2027+ has no company guidance to calibrate against, so the accrual" -Italic
SetC $ws 49 1 "methodology is retained there. The FY2026 adjustment is shown explicitly on row 42 rather than buried in an input." -Italic
$ws.Columns.Item(1).ColumnWidth = 44
for ($c=2; $c -le 7; $c++) { $ws.Columns.Item($c).ColumnWidth = 13 }
$ws.Columns.Item(8).ColumnWidth = 44

# ===================== MODEL =====================
$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
$ws.Name = "Model"
SetC $ws 1 1 "OPERATING MODEL, FREE CASH FLOW AND LIQUIDITY ROLL-FORWARD (GBP m)" -Bold
$ws.Cells.Item(1,1).Font.Size = 13
SetC $ws 2 1 '=Assumptions!C3' -Bold
for ($i=0; $i -lt 7; $i++) { SetC $ws 3 (2+$i) $yrs[$i] -Bold }

SectionHdr $ws 5 "OPERATING" 8
$oplab = @("Total wholesale volumes","Blended ASP (GBP k)","Revenue","Gross margin pct","Gross profit","Adjusted opex (ex D and A)","Adjusted EBITDA","Adjusted EBITDA margin","Adjusted D and A","Adjusted EBIT","Adjusted EBIT margin")
for ($i=0; $i -lt $oplab.Count; $i++) { SetC $ws (6+$i) 1 $oplab[$i] }
for ($i=0; $i -lt 7; $i++) {
  $c = [char](66+$i)
  SetC $ws 6  (2+$i) ('=Assumptions!{0}13' -f $c) -NumFmt $NUM0
  SetC $ws 7  (2+$i) ('=Assumptions!{0}14' -f $c) -NumFmt $NUM
  SetC $ws 8  (2+$i) ('={0}6*{0}7/1000' -f $c) -NumFmt $NUM -Bold
  SetC $ws 9  (2+$i) ('=Assumptions!{0}15' -f $c) -NumFmt $PCT
  SetC $ws 10 (2+$i) ('={0}8*{0}9' -f $c) -NumFmt $NUM
  SetC $ws 11 (2+$i) ('=-Assumptions!{0}16' -f $c) -NumFmt $NUM
  SetC $ws 12 (2+$i) ('={0}10+{0}11' -f $c) -NumFmt $NUM -Bold
  SetC $ws 13 (2+$i) ('={0}12/{0}8' -f $c) -NumFmt $PCT
  SetC $ws 14 (2+$i) ('=-Assumptions!{0}17' -f $c) -NumFmt $NUM
  SetC $ws 15 (2+$i) ('={0}12+{0}14' -f $c) -NumFmt $NUM -Bold
  SetC $ws 16 (2+$i) ('={0}15/{0}8' -f $c) -NumFmt $PCT
}

SectionHdr $ws 18 "FREE CASH FLOW" 8
$cflab = @("Adjusted EBITDA","Working capital movement","Lease payments","Cash tax","Other cash items","Net cash interest paid","Capex","FREE CASH FLOW")
for ($i=0; $i -lt $cflab.Count; $i++) { SetC $ws (19+$i) 1 $cflab[$i] }
for ($i=0; $i -lt 7; $i++) {
  $c = [char](66+$i)
  SetC $ws 19 (2+$i) ('={0}12' -f $c) -NumFmt $NUM
  SetC $ws 20 (2+$i) ('=Assumptions!{0}19' -f $c) -NumFmt $NUM
  SetC $ws 21 (2+$i) ('=Assumptions!{0}20' -f $c) -NumFmt $NUM
  SetC $ws 22 (2+$i) ('=Assumptions!{0}21' -f $c) -NumFmt $NUM
  SetC $ws 23 (2+$i) ('=Assumptions!{0}22' -f $c) -NumFmt $NUM
  SetC $ws 25 (2+$i) ('=-Assumptions!{0}18' -f $c) -NumFmt $NUM
  SetC $ws 26 (2+$i) ('=SUM({0}19:{0}25)' -f $c) -NumFmt $NUM -Bold
}
SetC $ws 24 2 -105.0 -NumFmt $NUM -FontColor $AMBER
for ($i=1; $i -lt 7; $i++) { $d = [char](65+$i); SetC $ws 24 (2+$i) ('=-Debt!{0}43' -f $d) -NumFmt $NUM }
SetC $ws 26 2 -410.0 -NumFmt $NUM -FontColor $BLUE -Bold
SetC $ws 27 1 "FY2025A free cash outflow is the reported figure [AR25], not a build. FY25 cash interest is PROVISIONAL." -Italic

SectionHdr $ws 29 "NET DEBT AND LIQUIDITY" 8
$dlab = @("Opening cash","Free cash flow","Net transaction proceeds","Other financing / debt movements","2029 notes repayment","2029 refinancing proceeds","Closing cash (before further funding)","Available facilities","TOTAL LIQUIDITY (before further funding)","Memo: undrawn DDTL (not in liquidity)","Gross debt (nominal)","Net debt","Net debt / adjusted EBITDA","Adjusted EBITDA / net cash interest","Liquidity headroom vs covenant")
for ($i=0; $i -lt $dlab.Count; $i++) { SetC $ws (30+$i) 1 $dlab[$i] }
SetC $ws 30 2 249.9 -NumFmt $NUM -FontColor $BLUE
SetC $ws 36 2 249.9 -NumFmt $NUM -FontColor $BLUE
SetC $ws 38 2 250.0 -NumFmt $NUM -FontColor $BLUE
SetC $ws 30 3 249.9 -NumFmt $NUM -FontColor $BLUE
for ($i=2; $i -lt 7; $i++) { $p=[char](65+$i); SetC $ws 30 (2+$i) ('={0}36' -f $p) -NumFmt $NUM }
for ($i=1; $i -lt 7; $i++) {
  $c = [char](66+$i)
  SetC $ws 31 (2+$i) ('={0}26' -f $c) -NumFmt $NUM
  SetC $ws 36 (2+$i) ('=SUM({0}30:{0}35)' -f $c) -NumFmt $NUM -Bold
  SetC $ws 37 (2+$i) 0 -NumFmt $NUM
  SetC $ws 38 (2+$i) ('={0}36+{0}37' -f $c) -NumFmt $NUM -Bold
  SetC $ws 39 (2+$i) '=(1-Assumptions!$B$34)*100' -NumFmt $NUM
  SetC $ws 41 (2+$i) ('={0}40-{0}36-1.5' -f $c) -NumFmt $NUM -Bold
  SetC $ws 42 (2+$i) ('=IF({0}12>0,{0}41/{0}12,"n.m.")' -f $c) -NumFmt $XX
  SetC $ws 43 (2+$i) ('=IF(-{0}24>0,{0}12/-{0}24,"n.m.")' -f $c) -NumFmt $XX
  SetC $ws 44 (2+$i) ('={0}38-Assumptions!$B$36' -f $c) -NumFmt $NUM
}
SetC $ws 32 3 '=450+Assumptions!$B$34*100-163-20-Assumptions!$B$33' -NumFmt $NUM
SetC $ws 33 3 60.8 -NumFmt $NUM -FontColor $BLUE
SetC $ws 33 8 "H1-26A financing and other investing inflows [H1]" -Italic
for ($i=2; $i -lt 7; $i++) { SetC $ws 32 (2+$i) 0 -NumFmt $NUM; SetC $ws 33 (2+$i) 0 -NumFmt $NUM }
for ($i=1; $i -lt 7; $i++) {
  if ($i -eq 4) {
    SetC $ws 34 (2+$i) -1357.3 -NumFmt $NUM
    SetC $ws 35 (2+$i) '=Assumptions!$B$35*1357.3' -NumFmt $NUM
  } else {
    SetC $ws 34 (2+$i) 0 -NumFmt $NUM
    SetC $ws 35 (2+$i) 0 -NumFmt $NUM
  }
}
SetC $ws 40 3 '=Debt!$D$14' -NumFmt $NUM
for ($i=2; $i -lt 7; $i++) { $c=[char](66+$i); $p=[char](65+$i); SetC $ws 40 (2+$i) ('={0}40+{1}34+{1}35' -f $p,$c) -NumFmt $NUM }
SetC $ws 45 1 "FUNDING GAP (further funding required)" -Bold
for ($i=1; $i -lt 7; $i++) { $c=[char](66+$i); SetC $ws 45 (2+$i) ('=-MIN(0,{0}38)' -f $c) -NumFmt $NUM -Bold -FontColor $RED }
SetC $ws 47 1 "FY2026 opening cash = 31-Dec-25 actual GBP 249.9m [H1]. Transaction modelled in FY2026." -Italic
SetC $ws 48 1 "Rows 36/38 are shown BEFORE any further funding. A negative figure quantifies the funding need - it is not a forecast of negative cash." -Italic -FontColor $RED
SetC $ws 49 1 "FY2026 modelled cash interest is c.GBP 12m above the guided c.GBP 160m: the model accrues a full year of note coupons, guidance reflects actual cash payment dates." -Italic -FontColor $AMBER
$ws.Columns.Item(1).ColumnWidth = 42
for ($c=2; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 12 }

# ===================== OUTPUTS =====================
$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
$ws.Name = "Outputs"
SetC $ws 1 1 "CREDIT METRICS DASHBOARD" -Bold
$ws.Cells.Item(1,1).Font.Size = 13
SetC $ws 2 1 '=Assumptions!C3' -Bold
SetC $ws 3 1 "PROVISIONAL - FY2027-31 drivers are unsourced judgement awaiting sign-off." -FontColor $RED -Bold
SetC $ws 4 1 "CAVEAT TRAVELS WITH THE NUMBER: the 20.16pct refi rate is a market-implied DISTRESS yield (embeds default probability), NOT a forecast new-issue coupon. Carry this into the memo and deck." -FontColor $RED -Bold
for ($i=0; $i -lt 7; $i++) { SetC $ws 5 (2+$i) $yrs[$i] -Bold }
$olab = @("Revenue","Adjusted EBITDA","Adjusted EBIT","Free cash flow","Net cash interest","Gross debt (nominal)","Net debt","Net debt / adj EBITDA","Adj EBITDA / cash interest","Total liquidity","Liquidity headroom vs covenant","Memo: undrawn DDTL")
$orow = @(8,12,15,26,24,40,41,42,43,38,44,39)
for ($i=0; $i -lt $olab.Count; $i++) {
  SetC $ws (6+$i) 1 $olab[$i]
  for ($j=0; $j -lt 7; $j++) {
    $c = [char](66+$j)
    if ($i -eq 7 -or $i -eq 8) { $fmt = $XX } else { $fmt = $NUM }
    SetC $ws (6+$i) (2+$j) ('=Model!{0}{1}' -f $c,$orow[$i]) -NumFmt $fmt
  }
}
SetC $ws 19 1 "HEADLINE ANSWER" -Bold
SetC $ws 20 1 "First year liquidity falls below the covenant threshold"
SetC $ws 20 2 '=IFERROR(INDEX($C$5:$H$5,MATCH(TRUE,INDEX($C$16:$H$16<0,0),0)),"No breach through FY2031")' -Bold
SetC $ws 21 1 "First year liquidity falls below zero"
SetC $ws 21 2 '=IFERROR(INDEX($C$5:$H$5,MATCH(TRUE,INDEX($C$15:$H$15<0,0),0)),"Positive through FY2031")' -Bold
SetC $ws 22 1 "2029 refinancing requirement (GBP m)"
SetC $ws 22 2 1357.3 -NumFmt $NUM -Bold
SetC $ws 23 1 "FY2027 cash interest step-up vs FY2026 (increase in cost)"
SetC $ws 23 2 '=-(Model!D24-Model!C24)' -NumFmt $NUM -Bold
SetC $ws 24 1 "Cumulative FCF FY2027-FY2031"
SetC $ws 24 2 '=SUM(Model!D26:H26)' -NumFmt $NUM -Bold
SetC $ws 25 1 "Peak funding gap by FY2031 (GBP m)"
SetC $ws 25 2 '=MAX(Model!C45:H45)' -NumFmt $NUM -Bold -FontColor $RED
SetC $ws 26 1 "Pro forma liquidity at 30-Jun-26 (disclosed)"
SetC $ws 26 2 '=Debt!D24' -NumFmt $NUM -Bold

SetC $ws 27 1 "Liquidity runway from 22-Jul-2026 (see Refinancing tab)"
SetC $ws 27 2 "BASE c.19 months (crosses zero c.Feb-2028)  |  DOWNSIDE c.16 months (crosses zero c.Nov-2027)" -Bold

SectionHdr $ws 28 "COVENANT THRESHOLD SENSITIVITY - the breach date is a function of an UNSOURCED input" 8
SetC $ws 29 1 "Minimum liquidity threshold (GBP m)" -Bold
SetC $ws 29 2 "First breach year" -Bold
SetC $ws 29 3 "Headroom at FY2027 (GBP m)" -Bold
$thr = @(75,100,125)
for ($i=0; $i -lt 3; $i++) {
  $r = 30 + $i
  SetC $ws $r 1 $thr[$i] -NumFmt $NUM -FontColor $RED
  SetC $ws $r 2 ('=IFERROR(INDEX($C$5:$H$5,MATCH(TRUE,INDEX(Model!$C$38:$H$38<A{0},0),0)),"No breach to FY2031")' -f $r) -Bold
  SetC $ws $r 3 ('=Model!D38-A{0}' -f $r) -NumFmt $NUM
}
SetC $ws 34 1 "The SSTL minimum-liquidity covenant is disclosed [H1] but its LEVEL is not. The model uses Assumptions!B36." -Italic -FontColor $RED
SetC $ws 35 1 "Across the plausible 75-125 range the breach year is unchanged, so the conclusion does not hinge on this input." -Italic
SetC $ws 36 1 "This block reflects the ACTIVE scenario (Assumptions!B3)." -Italic

SectionHdr $ws 38 "SCENARIO COMPARISON - static snapshot taken 2026-08-07, refresh if drivers change" 8
SetC $ws 39 1 "Metric" -Bold
SetC $ws 39 2 "BASE" -Bold
SetC $ws 39 3 "DOWNSIDE" -Bold
SetC $ws 39 4 "Delta" -Bold
$cmp = @(
 @(40,"FY2027 revenue (GBP m)",           "1,671.4","1,357.8","-313.6  (-18.8pct)"),
 @(41,"FY2027 adjusted EBITDA (GBP m)",   "288.3",  "129.5",  "-158.8  (-55pct)"),
 @(42,"FY2027 EBITDA margin",             "17.3pct","9.5pct", "-780bps"),
 @(43,"FY2027 free cash flow (GBP m)",    "-277.6", "-455.4", "-177.8"),
 @(44,"Covenant breach",                  "FY2027", "FY2027", "same"),
 @(45,"Liquidity exhausted (below zero)", "FY2028", "FY2027", "one year earlier"),
 @(46,"Peak funding gap (GBP m)",         "1,318.8","2,162.5","+843.7"),
 @(47,"FY2031 net debt / EBITDA",         "10.2x",  "29.9x",  "+19.7x"),
 @(48,"FY2029 EBITDA / cash interest",    "1.0x",   "0.4x",   "-0.6x"),
 @(49,"DDTL",                             "undrawn","drawn",  "+GBP 100m debt and liquidity")
)
foreach ($c in $cmp) { for ($i=0; $i -lt 4; $i++) { SetC $ws $c[0] (1+$i) $c[1+$i] } }
$ws.Range($ws.Cells.Item(45,1), $ws.Cells.Item(45,4)).Font.Bold = $true
SetC $ws 51 1 "Downside = single combined stress (volumes -23pct, gross margin 32.0pct, SONIA 5.25pct) applied together from FY2027," -Italic
SetC $ws 52 1 "with capex and opex deliberately held flat - a NO-RESPONSE case matching management's own reverse stress test." -Italic
SetC $ws 53 1 "Even in the base case the capital structure does not work. The downside brings the wall forward by a year and doubles the gap." -Bold
$ws.Columns.Item(1).ColumnWidth = 46
for ($c=2; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 13 }

# ===================== YTM =====================
$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
$ws.Name = "YTM"
SetC $ws 1 1 "IMPLIED YIELD TO MATURITY ON THE 2029 NOTES - basis for the refinancing rate" -Bold
$ws.Cells.Item(1,1).Font.Size = 13
SetC $ws 2 1 "Settlement 30-Jun-2026. Maturity 31-Mar-2029 (2.75 years). Semi-annual coupons, 6 payments remaining." -Italic
SetC $ws 3 1 "Prices are the disclosed fair values at 30-Jun-2026 [H1 fair value note] - see SOURCES.md section 3." -Italic

SetC $ws 5 1 "Tranche" -Bold
SetC $ws 5 2 "Nominal" -Bold
SetC $ws 5 3 "Fair value" -Bold
SetC $ws 5 4 "Clean price pct" -Bold
SetC $ws 5 5 "Coupon pct" -Bold
SetC $ws 5 6 "Accrued pct" -Bold
SetC $ws 5 7 "Dirty price pct" -Bold
SetC $ws 5 8 "Implied YTM pct" -Bold
$yt = @(
 @(6,"USD SSN 10.000pct",   792.3, 629.9, 10.000, 20.02),
 @(7,"GBP SSN 10.375pct (465m)", 465.0, 372.3, 10.375, 20.14),
 @(8,"GBP SSN 10.375pct (100m)", 100.0, 77.8,  10.375, 21.44),
 @(9,"BLENDED",            1357.3, 1080.0, 10.156, 20.16)
)
foreach ($y in $yt) {
  $r = $y[0]
  SetC $ws $r 1 $y[1]
  SetC $ws $r 2 $y[2] -NumFmt $NUM -FontColor $BLUE
  SetC $ws $r 3 $y[3] -NumFmt $NUM -FontColor $BLUE
  SetC $ws $r 4 ('=100*C{0}/B{0}' -f $r) -NumFmt '0.00'
  SetC $ws $r 5 $y[4] -NumFmt '0.000' -FontColor $BLUE
  SetC $ws $r 6 ('=E{0}*0.25' -f $r) -NumFmt '0.00'
  SetC $ws $r 7 ('=D{0}+F{0}' -f $r) -NumFmt '0.00'
  SetC $ws $r 8 $y[5] -NumFmt '0.00' -Bold
}
SetC $ws 11 1 "Method" -Bold
SetC $ws 12 1 "Dirty price = clean price + accrued interest (3 months since the 31-Mar coupon)."
SetC $ws 13 1 "Solve for y such that: sum over t of (coupon/2)/(1+y/2)^(2t) + 100/(1+y/2)^(2*2.75) = dirty price,"
SetC $ws 14 1 "with t = 0.25, 0.75, 1.25, 1.75, 2.25, 2.75 years. Solved by bisection to 1e-6."
SetC $ws 15 1 "Blended coupon = (792.3 x 10.000 + 565.0 x 10.375) / 1357.3 = 10.156 pct."
SetC $ws 17 1 "RESULT: blended implied YTM = 20.16 pct. This is the refinancing rate used on the Assumptions tab." -Bold
SetC $ws 19 1 "INTERPRETIVE CAVEAT - read before quoting this number:" -Bold -FontColor $RED
SetC $ws 20 1 "20.16 pct is a DISTRESS yield. It embeds the market's probability of default, not a clean forecast of a new-issue coupon." -FontColor $RED
SetC $ws 21 1 "Using it as the 2029 refinancing rate is deliberately conservative: it prices the refinancing at what the market thinks" -FontColor $RED
SetC $ws 22 1 "this risk costs TODAY. If AML reaches 2029 in a position to refinance at all, the credit would likely have improved and" -FontColor $RED
SetC $ws 23 1 "the true coupon would be lower. The honest reading is not 'AML will pay 20 pct' but 'the market does not currently believe" -FontColor $RED
SetC $ws 24 1 "these notes can be refinanced on economic terms' - which is itself the finding." -FontColor $RED
$ws.Columns.Item(1).ColumnWidth = 30
for ($c=2; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 15 }

# ===================== ATTRIBUTION =====================
$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
$ws.Name = "Attribution"
SetC $ws 1 1 "WHAT DRIVES THE LEVERAGE AND COVERAGE REVERSAL - stepwise attribution" -Bold
$ws.Cells.Item(1,1).Font.Size = 13
SetC $ws 2 1 "Cumulative walk from the pre-review base case to the current one. Each step adds one change." -Italic
SetC $ws 3 1 "Produced by re-running this workbook with each input reverted; S4 reconciles exactly to the live model." -Italic

SetC $ws 5 1 "Step" -Bold
SetC $ws 5 2 "FY29 interest" -Bold
SetC $ws 5 3 "FY31 interest" -Bold
SetC $ws 5 4 "FY29 leverage" -Bold
SetC $ws 5 5 "FY31 leverage" -Bold
SetC $ws 5 6 "FY29 coverage" -Bold
SetC $ws 5 7 "FY31 coverage" -Bold
SetC $ws 5 8 "Peak funding gap" -Bold
$att = @(
 @(6,"S0  Pre-review base case",              200.8, 203.7, 5.71, 5.56, 1.99, 2.21, 560.0),
 @(7,"S1  + margin cap to 17.1pct",           200.8, 203.7, 8.04, 9.17, 1.52, 1.57, 995.3),
 @(8,"S2  + SONIA 4.00 to 3.72pct",           199.6, 202.4, 8.03, 9.15, 1.53, 1.58, 988.4),
 @(9,"S3  + FY2026 interest calibration",     199.6, 202.4, 7.99, 9.12, 1.53, 1.58, 976.9),
 @(10,"S4 + refi 11.0 to 20.16pct (CURRENT)", 292.8, 326.7, 8.29, 10.18, 1.05, 0.98, 1318.8)
)
foreach ($s in $att) {
  $r = $s[0]
  SetC $ws $r 1 $s[1]
  SetC $ws $r 2 $s[2] -NumFmt $NUM
  SetC $ws $r 3 $s[3] -NumFmt $NUM
  SetC $ws $r 4 $s[4] -NumFmt $XX
  SetC $ws $r 5 $s[5] -NumFmt $XX
  SetC $ws $r 6 $s[6] -NumFmt $XX
  SetC $ws $r 7 $s[7] -NumFmt $XX
  SetC $ws $r 8 $s[8] -NumFmt $NUM
}
$ws.Range($ws.Cells.Item(10,1), $ws.Cells.Item(10,8)).Font.Bold = $true

SetC $ws 12 1 "CONTRIBUTION TO THE TOTAL SWING" -Bold
SetC $ws 13 1 "Metric" -Bold
SetC $ws 13 2 "Total swing" -Bold
SetC $ws 13 3 "Margin cap" -Bold
SetC $ws 13 4 "SONIA" -Bold
SetC $ws 13 5 "FY26 calib." -Bold
SetC $ws 13 6 "Refi rate" -Bold
$con = @(
 @(14,"FY2031 leverage",       "+4.62x", "+3.61x  (78pct)", "-0.02x  (0pct)",  "-0.03x  (-1pct)", "+1.06x  (23pct)"),
 @(15,"FY2029 coverage",       "-0.94x", "-0.47x  (50pct)", "+0.01x  (-1pct)", "0.00x   (0pct)",  "-0.48x  (51pct)"),
 @(16,"FY2031 coverage",       "-1.23x", "-0.64x  (52pct)", "+0.01x  (-1pct)", "0.00x   (0pct)",  "-0.60x  (49pct)"),
 @(17,"Peak funding gap (GBP m)","+758.8","+435.3  (57pct)", "-6.9   (-1pct)",  "-11.5   (-2pct)", "+341.9  (45pct)")
)
foreach ($c in $con) { for ($i=0; $i -lt 6; $i++) { SetC $ws $c[0] (1+$i) $c[1+$i] } }

SetC $ws 19 1 "CONCLUSION" -Bold
SetC $ws 20 1 "The leverage reversal is driven MAINLY by the margin cap (78pct of the FY2031 swing), not the refi rate (23pct)." -Bold
SetC $ws 21 1 "The coverage collapse is roughly an even split between the two (c.50/50)."
SetC $ws 22 1 "SONIA and the FY2026 calibration are immaterial (under 2pct each) and both slightly REDUCE stress."
SetC $ws 24 1 "This is NOT a compounding of unexamined provisional assumptions. Two deliberate corrections did the work, and both"
SetC $ws 25 1 "replaced a weaker input with a better-grounded one: the margin cap is anchored to AML's own FY2024 high (sourced),"
SetC $ws 26 1 "and the refi rate is computed from the disclosed market price (sourced) rather than guessed."
SetC $ws 28 1 "ROBUSTNESS: step S1 shows that with the margin cap alone - refi still at 11pct - leverage already rises to 9.17x and" -Bold
SetC $ws 29 1 "the funding gap reaches GBP 995m. The conclusion does NOT depend on the distress yield." -Bold
$ws.Columns.Item(1).ColumnWidth = 42
for ($c=2; $c -le 8; $c++) { $ws.Columns.Item($c).ColumnWidth = 17 }

# ===================== REFINANCING =====================
$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
$ws.Name = "Refinancing"
SetC $ws 1 1 "2029 REFINANCING FEASIBILITY AND THE SELF-SUSTAINING TEST" -Bold
$ws.Cells.Item(1,1).Font.Size = 13
SetC $ws 2 1 "Answers two of the brief's core questions: can AML refinance the 2029 notes, and when does it become self-sustaining?" -Italic
SetC $ws 3 1 "All figures BASE CASE. GBP 1,357.3m of senior secured notes mature March 2029." -Italic

SectionHdr $ws 5 "A. WHAT WOULD REFINANCING COST? (FY2030, first full year post-refinancing)" 6
SetC $ws 6 1 "Refinancing rate" -Bold
SetC $ws 6 2 "FY30 cash interest" -Bold
SetC $ws 6 3 "FY30 coverage" -Bold
SetC $ws 6 4 "FY30 free cash flow" -Bold
SetC $ws 6 5 "Peak funding gap" -Bold
SetC $ws 6 6 "Comment" -Bold
$rf = @(
 @(7, "0.00pct (hypothetical)", 53.1, 5.89, -115.1, 566.3, "Even at ZERO cost of debt, free cash flow stays negative"),
 @(8, "7.50pct",               154.9, 2.02, -216.9, 846.2, "Coverage reaches 2.0x only here - well inside investment-grade pricing"),
 @(9, "10.375pct (existing coupon)",193.9,1.61,-255.9, 953.5, "Refinancing at the CURRENT coupon still leaves 1.6x and negative FCF"),
 @(10,"12.50pct",              222.8, 1.41, -284.8, 1032.8, ""),
 @(11,"15.00pct",              256.7, 1.22, -318.7, 1126.2, ""),
 @(12,"20.16pct (market-implied)",326.7,0.96,-388.7,1318.8, "DISTRESS yield, not a forecast coupon - see YTM tab. Coverage below 1.0x")
)
foreach ($r in $rf) { SetC $ws $r[0] 1 $r[1]; SetC $ws $r[0] 2 $r[2] -NumFmt $NUM; SetC $ws $r[0] 3 $r[3] -NumFmt $XX; SetC $ws $r[0] 4 $r[4] -NumFmt $NUM; SetC $ws $r[0] 5 $r[5] -NumFmt $NUM; SetC $ws $r[0] 6 $r[6] -Italic }
$ws.Range($ws.Cells.Item(7,1), $ws.Cells.Item(7,6)).Font.Bold = $true
$ws.Range($ws.Cells.Item(9,1), $ws.Cells.Item(9,6)).Font.Bold = $true

SetC $ws 14 1 "THE DECISIVE RESULT" -Bold
SetC $ws 15 1 "At a 0pct refinancing cost - an impossible best case - FY2030 free cash flow is still MINUS GBP 115m." -Bold -FontColor $RED
SetC $ws 16 1 "Refinancing terms are therefore not the binding constraint. No achievable coupon makes this capital structure work." -Bold -FontColor $RED
SetC $ws 17 1 "Refinancing feasibility is a question about the BUSINESS, not about credit markets."

SectionHdr $ws 19 "B. SELF-SUSTAINING TEST - can the business fund its own capex, before any interest?" 6
SetC $ws 20 1 "Base case (GBP m)" -Bold
for ($i=1; $i -lt 7; $i++) { SetC $ws 20 (1+$i) $yrs[$i] -Bold }
$ssl = @("Adjusted EBITDA","Capex","EBITDA less capex","Cash interest","Free cash flow")
for ($i=0; $i -lt 5; $i++) { SetC $ws (21+$i) 1 $ssl[$i] }
$ssv = @(
 @(264.6, 288.3, 299.5, 306.2, 313.0, 320.1),
 @(-300.0,-350.0,-350.0,-350.0,-350.0,-300.0),
 @(-35.4, -61.7, -50.5, -43.8, -37.0, 20.1),
 @(-160.0,-191.0,-191.0,-292.8,-326.7,-326.7),
 @(-230.4,-277.6,-266.5,-361.6,-388.7,-331.6)
)
for ($r=0; $r -lt 5; $r++) { for ($i=0; $i -lt 6; $i++) { SetC $ws (21+$r) (2+$i) $ssv[$r][$i] -NumFmt $NUM } }
$ws.Range($ws.Cells.Item(23,1), $ws.Cells.Item(23,7)).Font.Bold = $true

SetC $ws 27 1 "EBITDA does not cover CAPEX ALONE until FY2031 - and then by only GBP 20m, before a penny of interest." -Bold -FontColor $RED
SetC $ws 28 1 "On these drivers Aston Martin does NOT become self-sustaining within the forecast horizon." -Bold -FontColor $RED

SectionHdr $ws 30 "C. WHAT WOULD HAVE TO BE TRUE for FY2031 free cash flow to reach zero?" 6
$wt = @(
 @(31,"Required EBITDA",              "GBP 651.7m", "vs GBP 320.1m modelled - more than DOUBLE"),
 @(32,"Implied revenue (at 35.5pct GM, opex GBP 350m)","GBP 2,823m","vs GBP 1,888m modelled - c.+50pct"),
 @(33,"Implied wholesale volumes (at ASP GBP 286k)","c.9,870 units","vs 6,600 modelled, and vs 5,448 actual in FY2025"),
 @(34,"Alternative: refinance at existing 10.375pct","still needs EBITDA of c.GBP 519m","c.+62pct vs modelled")
)
foreach ($w in $wt) { SetC $ws $w[0] 1 $w[1] -Bold; SetC $ws $w[0] 2 $w[2] -Bold; SetC $ws $w[0] 3 $w[3] -Italic }
SetC $ws 36 1 "For context, FY2025 actual wholesale volumes were 5,448 and peak historic adjusted EBITDA was GBP 271m (FY2024)." -Italic

SectionHdr $ws 38 "D. LIQUIDITY RUNWAY from the 22-Jul-2026 financing" 6
SetC $ws 39 1 "Scenario" -Bold
SetC $ws 39 2 "Crosses zero" -Bold
SetC $ws 39 3 "Runway" -Bold
SetC $ws 39 4 "Basis" -Bold
SetC $ws 40 1 "BASE"
SetC $ws 40 2 "c. Feb-2028"
SetC $ws 40 3 "c. 19 months"
SetC $ws 40 4 "FY2028 opening liquidity GBP 29.7m against FY2028 free cash outflow of GBP 266.5m (11.1pct through the year)" -Italic
SetC $ws 41 1 "DOWNSIDE"
SetC $ws 41 2 "c. Nov-2027"
SetC $ws 41 3 "c. 16 months"
SetC $ws 41 4 "FY2027 opening liquidity GBP 407.3m (incl. drawn DDTL) against FY2027 free cash outflow of GBP 455.4m (89.4pct through)" -Italic
$ws.Range($ws.Cells.Item(40,1), $ws.Cells.Item(41,3)).Font.Bold = $true
SetC $ws 43 1 "Straight-line interpolation within the year - the model is annual, so treat these as indicative quarters, not dates." -Italic -FontColor $RED

SectionHdr $ws 45 "E. ANSWERS TO THE BRIEF'S QUESTIONS" 6
$ans = @(
 @(46,"Can AML afford the additional interest?","No. FY2027 cash interest of GBP 191m against EBITDA of GBP 288m leaves 1.5x coverage, and free cash flow stays negative throughout."),
 @(47,"Can it refinance the 2029 notes?","Not on terms that help. Even at a 0pct coupon FY2030 free cash flow is minus GBP 115m. The constraint is the business, not the market."),
 @(48,"Is further funding required before 2031?","Yes. Base case peak funding gap GBP 1,319m; downside GBP 2,163m. Required from FY2028 (base) or FY2027 (downside)."),
 @(49,"When does it become self-sustaining?","Not within the horizon. EBITDA does not cover capex alone until FY2031, and never covers capex plus interest."),
 @(50,"Is further restructuring or a capital raise likely?","Yes - on these drivers it is close to unavoidable well before the March 2029 maturity.")
)
foreach ($x in $ans) { SetC $ws $x[0] 1 $x[1] -Bold; SetC $ws $x[0] 2 $x[2] }
$ws.Columns.Item(1).ColumnWidth = 46
$ws.Columns.Item(2).ColumnWidth = 34
for ($c=3; $c -le 6; $c++) { $ws.Columns.Item($c).ColumnWidth = 22 }

# ===================== RECONCILIATION =====================
$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
$ws.Name = "Reconciliation"
SetC $ws 1 1 "WHY THE MODEL EXHAUSTS LIQUIDITY AT A MILDER SHOCK THAN MANAGEMENT'S REVERSE STRESS TEST" -Bold
$ws.Cells.Item(1,1).Font.Size = 13
SetC $ws 2 1 "THE CHALLENGE: our downside uses -23pct total volumes (management's -25pct CORE) yet fully exhausts liquidity." -Italic
SetC $ws 3 1 "Management's disclosure implies -55pct core is needed to exhaust liquidity - roughly double our stress. Why?" -Italic

SectionHdr $ws 5 "TEST 1 - LEVER WALK: which stress lever actually causes exhaustion?" 6
SetC $ws 6 1 "Variant (all vs FY2027 year-end liquidity)" -Bold
SetC $ws 6 2 "FY27 EBITDA" -Bold
SetC $ws 6 3 "FY27 liquidity" -Bold
SetC $ws 6 4 "Change" -Bold
SetC $ws 6 5 "Liquidity exhausted" -Bold
$lw = @(
 @(7, "BASE (no stress)",              288.3,  29.7,  "-",       "FY2028"),
 @(8, "V1  core volumes -25pct ONLY",  176.8, -81.8,  "-111.5",  "FY2027"),
 @(9, "V2  + gross margin 32pct",      129.3, -129.3, "-47.5",   "FY2027"),
 @(10,"V3  + SONIA 5.25pct",           129.3, -136.2, "-6.9",    "FY2027"),
 @(11,"V4  + DDTL drawn (= DOWNSIDE)", 129.3, -48.2,  "+88.0",   "FY2027")
)
foreach ($l in $lw) { SetC $ws $l[0] 1 $l[1]; SetC $ws $l[0] 2 $l[2] -NumFmt $NUM; SetC $ws $l[0] 3 $l[3] -NumFmt $NUM; SetC $ws $l[0] 4 $l[4]; SetC $ws $l[0] 5 $l[5] }
$ws.Range($ws.Cells.Item(8,1), $ws.Cells.Item(8,5)).Font.Bold = $true
SetC $ws 13 1 "FINDING: the margin assumption is NOT the cause. A volume-only -25pct stress already exhausts liquidity" -Bold
SetC $ws 14 1 "in FY2027 (-GBP 81.8m). The margin cut adds a further -GBP 47.5m but is not required for exhaustion." -Bold
SetC $ws 15 1 "SONIA is trivial (-GBP 6.9m). Drawing the DDTL ADDS GBP 88.0m net of its own interest." -Italic

SectionHdr $ws 17 "TEST 2 - VOLUME-ONLY LADDER: what shock does our model actually need?" 6
SetC $ws 18 1 "Core volume cut (margin and SONIA at base, DDTL undrawn)" -Bold
SetC $ws 18 2 "FY27 revenue" -Bold
SetC $ws 18 3 "FY27 EBITDA" -Bold
SetC $ws 18 4 "FY27 liquidity" -Bold
SetC $ws 18 5 "Liquidity exhausted" -Bold
$lad = @(
 @(19,"0pct (base)",   1671.4, 288.3,  29.7, "FY2028"),
 @(20,"-5pct",         1608.6, 266.0,   7.4, "FY2028"),
 @(21,"-6.6pct",       1588.4, 258.9,   0.3, "FY2028  <- BREAKEVEN"),
 @(22,"-10pct",        1545.7, 243.7, -14.9, "FY2027"),
 @(23,"-25pct (mgmt covenant marker)", 1357.2, 176.8, -81.8, "FY2027"),
 @(24,"-40pct",        1168.6, 109.9,-148.8, "FY2027"),
 @(25,"-55pct (mgmt liquidity marker)", 980.1, 42.9,-215.7, "FY2027")
)
foreach ($l in $lad) { SetC $ws $l[0] 1 $l[1]; SetC $ws $l[0] 2 $l[2] -NumFmt $NUM; SetC $ws $l[0] 3 $l[3] -NumFmt $NUM; SetC $ws $l[0] 4 $l[4] -NumFmt $NUM; SetC $ws $l[0] 5 $l[5] }
$ws.Range($ws.Cells.Item(21,1), $ws.Cells.Item(21,5)).Font.Bold = $true
SetC $ws 27 1 "FINDING: our model tips into FY2027 exhaustion at only -6.6pct core volumes - versus management's -55pct." -Bold -FontColor $RED
SetC $ws 28 1 "The gap is therefore NOT in the stress levers. It is in the BASE CASE we are stressing from." -Bold -FontColor $RED

SectionHdr $ws 30 "WHY: the base case is already on the cliff edge" 6
SetC $ws 31 1 "Base FY2027 committed outflow vs earnings (GBP m)" -Bold
$bd = @(@(32,"Adjusted EBITDA",288.3),@(33,"Cash interest",-191.0),@(34,"Capex",-350.0),@(35,"Working capital, leases, tax",-25.0),@(36,"FREE CASH FLOW",-277.6))
foreach ($b in $bd) { SetC $ws $b[0] 1 $b[1]; SetC $ws $b[0] 2 $b[2] -NumFmt $NUM }
$ws.Range($ws.Cells.Item(36,1), $ws.Cells.Item(36,2)).Font.Bold = $true
SetC $ws 38 1 "Interest (GBP 191m) plus capex (GBP 350m) = GBP 541m of largely committed outflow against GBP 288m of EBITDA."
SetC $ws 39 1 "Base FY2027 year-end liquidity is only GBP 29.7m. Any material stress tips it negative - hence -6.6pct."

SectionHdr $ws 41 "TEST 3 - is this an artefact of OUR capex phasing? No." 6
SetC $ws 42 1 "Capex phasing FY2026-30 (total held at c.GBP 1.7bn, which IS sourced guidance)" -Bold
SetC $ws 42 2 "FY27 liquidity" -Bold
SetC $ws 42 3 "FY28 liquidity" -Bold
SetC $ws 42 4 "Liquidity exhausted" -Bold
$cx = @(
 @(43,"A  300/350/350/350/350 (current)", 29.7, -236.7, "FY2028"),
 @(44,"B  300/300/350/375/375",           79.7, -186.7, "FY2028"),
 @(45,"C  300/250/350/400/400",          129.7, -136.7, "FY2028")
)
foreach ($c in $cx) { SetC $ws $c[0] 1 $c[1]; SetC $ws $c[0] 2 $c[2] -NumFmt $NUM; SetC $ws $c[0] 3 $c[3] -NumFmt $NUM; SetC $ws $c[0] 4 $c[4] }
SetC $ws 47 1 "Phasing moves FY2027 liquidity by up to GBP 100m but does NOT change the exhaustion year. The programme TOTAL" -Bold
SetC $ws 48 1 "is sourced guidance, so re-phasing only shifts the burn between years - it cannot avoid it." -Bold

SectionHdr $ws 50 "THE FOUR REASONS THE TWO NUMBERS ARE NOT LIKE-FOR-LIKE" 6
$rz = @(
 @(51,"1. Different starting forecast","Management stresses THEIR forecast; we stress OUR reconstruction. Ours already burns GBP 278m in FY2027 and ends with GBP 29.7m. Theirs must be materially stronger - the -6.6pct vs -55pct gap MEASURES that divergence."),
 @(52,"2. Different horizon","Management's window ends 30-Sep-2027. Our FY2027 ends 31-Dec-2027 - roughly one extra quarter of burn (c.GBP 70-90m at our run-rate)."),
 @(53,"3. Different metric","Management tests volumes ONLY, holding per-unit economics flat. Our downside also cuts margin and raises SONIA - though Test 1 shows volumes alone are already sufficient."),
 @(54,"4. Mitigating actions","Both exclude them, so this is neutral - but management states it would act, and in practice would cut capex first.")
)
foreach ($r in $rz) { SetC $ws $r[0] 1 $r[1] -Bold; SetC $ws $r[0] 2 $r[2] }

SectionHdr $ws 56 "HOW TO DEFEND THIS IN INTERVIEW" 6
SetC $ws 57 1 "Do NOT claim the model contradicts management. It does not - it stresses a different, more conservative base."
SetC $ws 58 1 "The defensible statement: 'On a base case capped at AML's own historic best margin and carrying its own guided"
SetC $ws 59 1 "capex programme, the capital structure is already cash-negative before any stress. That is the finding. The"
SetC $ws 60 1 "downside then simply confirms there is no room for error.'"
SetC $ws 62 1 "HONEST CAVEAT: this makes the headline conclusion sensitive to the BASE case, not just the downside." -Bold -FontColor $RED
SetC $ws 63 1 "Management's own forecast is NOT disclosed, so the divergence cannot be reconciled directly - only bounded." -Bold -FontColor $RED
$ws.Columns.Item(1).ColumnWidth = 46
$ws.Columns.Item(2).ColumnWidth = 105
for ($c=3; $c -le 6; $c++) { $ws.Columns.Item($c).ColumnWidth = 20 }

# ===================== MATURITY =====================
$ws = $wb.Worksheets.Add([System.Reflection.Missing]::Value, $wb.Worksheets.Item($wb.Worksheets.Count))
$ws.Name = "Maturity"
SetC $ws 1 1 "DEBT MATURITY PROFILE (GBP m, nominal) - POST-TRANSACTION" -Bold
$ws.Cells.Item(1,1).Font.Size = 13
SetC $ws 3 1 "Instrument" -Bold
$my = @("2026","2027","2028","2029","2030","2031")
for ($i=0; $i -lt 6; $i++) { SetC $ws 3 (2+$i) $my[$i] -Bold }
$mat = @(
 @(4,"USD SSN (Mar-29)",    0,0,0,'=Debt!D4',0,0),
 @(5,"GBP SSN (Mar-29)",    0,0,0,'=Debt!D5+Debt!D6',0,0),
 @(6,"SSTL (Jul-31)",       0,0,0,0,0,'=Debt!D7'),
 @(7,"DDTL (Jul-31)",       0,0,0,0,0,'=Debt!D8'),
 @(8,"Inventory financing", '=Debt!D12',0,0,0,0,0),
 @(9,"Other bank loans",    '=Debt!D11',0,0,0,0,0)
)
foreach ($m in $mat) {
  SetC $ws $m[0] 1 $m[1]
  for ($i=0; $i -lt 6; $i++) { SetC $ws $m[0] (2+$i) $m[2+$i] -NumFmt $NUM }
}
SetC $ws 10 1 "TOTAL" -Bold
for ($i=0; $i -lt 6; $i++) { $c=[char](66+$i); SetC $ws 10 (2+$i) ('=SUM({0}4:{0}9)' -f $c) -NumFmt $NUM -Bold }
SetC $ws 12 1 "Lease liabilities (GBP 89.7m) amortise across the period and are excluded above." -Italic
SetC $ws 13 1 "The Mar-2029 wall is GBP 1,357.3m nominal - c.4x the pro forma liquidity of c.GBP 340m." -Bold
$ws.Columns.Item(1).ColumnWidth = 34
for ($c=2; $c -le 7; $c++) { $ws.Columns.Item($c).ColumnWidth = 13 }

$wb.Worksheets.Item("README").Activate()
$wb.SaveAs($out, 51)
$wb.Close($false)
$xl.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
"SAVED: $out"
