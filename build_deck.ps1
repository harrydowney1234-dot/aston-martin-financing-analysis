$ErrorActionPreference='Stop'
$P = [char]0x00A3   # pound sign, injected to keep this script pure ASCII
$out = "C:\Users\harry\aston-martin-financing-analysis\AML_Financing_Deck.pptx"
if (Test-Path $out) { Remove-Item $out -Force }

$NAVY  = 4401680     # RGB(16,42,67)
$RED   = 192         # RGB(192,0,0)
$GREY  = 6579300     # RGB(100,100,100)
$LIGHT = 15921906    # RGB(242,242,242)
$WHITE = 16777215

$pp = New-Object -ComObject PowerPoint.Application
$pres = $pp.Presentations.Add(0)
$pres.PageSetup.SlideWidth = 960
$pres.PageSetup.SlideHeight = 540

function NewSlide($title, $headline) {
  $s = $pres.Slides.Add($pres.Slides.Count + 1, 12)   # ppLayoutBlank
  $t = $s.Shapes.AddTextbox(1, 40, 28, 880, 34)
  $t.TextFrame.TextRange.Text = $title
  $t.TextFrame.TextRange.Font.Size = 26
  $t.TextFrame.TextRange.Font.Bold = -1
  $t.TextFrame.TextRange.Font.Color.RGB = $NAVY
  $r = $s.Shapes.AddLine(40, 66, 920, 66)
  $r.Line.ForeColor.RGB = $NAVY
  $r.Line.Weight = 1.5
  $h = $s.Shapes.AddTextbox(1, 40, 72, 880, 30)
  $h.TextFrame.TextRange.Text = $headline
  $h.TextFrame.TextRange.Font.Size = 15
  $h.TextFrame.TextRange.Font.Italic = -1
  $h.TextFrame.TextRange.Font.Color.RGB = $GREY
  return $s
}
function Txt($s, $l, $t, $w, $h, $text, $size, $bold, $color) {
  $b = $s.Shapes.AddTextbox(1, $l, $t, $w, $h)
  $b.TextFrame.TextRange.Text = $text
  $b.TextFrame.TextRange.Font.Size = $size
  if ($bold) { $b.TextFrame.TextRange.Font.Bold = -1 }
  if ($color -ne $null) { $b.TextFrame.TextRange.Font.Color.RGB = $color }
  $b.TextFrame.WordWrap = -1
  return $b
}
function Tbl($s, $rows, $cols, $l, $t, $w, $h, $data, $widths, $leftText) {
  $sh = $s.Shapes.AddTable($rows, $cols, $l, $t, $w, $h)
  $tb = $sh.Table
  for ($r=1; $r -le $rows; $r++) {
    for ($c=1; $c -le $cols; $c++) {
      $cell = $tb.Cell($r,$c)
      $cell.Shape.TextFrame.TextRange.Text = [string]$data[$r-1][$c-1]
      $cell.Shape.TextFrame.TextRange.Font.Size = 12
      $cell.Shape.TextFrame.MarginTop = 3; $cell.Shape.TextFrame.MarginBottom = 3
      if ($r -eq 1) {
        $cell.Shape.Fill.ForeColor.RGB = $NAVY
        $cell.Shape.TextFrame.TextRange.Font.Color.RGB = $WHITE
        $cell.Shape.TextFrame.TextRange.Font.Bold = -1
      } else {
        $cell.Shape.Fill.ForeColor.RGB = $WHITE
        $cell.Shape.TextFrame.TextRange.Font.Color.RGB = 0
      }
      if ($c -gt 1 -and -not $leftText) { $cell.Shape.TextFrame.TextRange.ParagraphFormat.Alignment = 3 }
    }
  }
  if ($widths -ne $null) { for ($c=1; $c -le $cols; $c++) { $tb.Columns($c).Width = $widths[$c-1] } }
  return $sh
}
function CalloutBox($s, $l, $t, $w, $h, $text, $color) {
  $b = $s.Shapes.AddShape(1, $l, $t, $w, $h)   # rectangle
  $b.Fill.ForeColor.RGB = $LIGHT
  $b.Line.ForeColor.RGB = $color
  $b.Line.Weight = 1
  $b.TextFrame.TextRange.Text = $text
  $b.TextFrame.TextRange.Font.Size = 11
  $b.TextFrame.TextRange.Font.Italic = -1
  $b.TextFrame.TextRange.Font.Color.RGB = $color
  $b.TextFrame.TextRange.ParagraphFormat.Alignment = 1
  $b.TextFrame.MarginLeft = 10; $b.TextFrame.MarginRight = 10
  return $b
}
function Notes($s, $text) { $s.NotesPage.Shapes.Item(2).TextFrame.TextRange.Text = $text }

# ---------------- SLIDE 1 ----------------
$s = NewSlide "Executive conclusion" "Aston Martin Lagonda | $($P)550m private credit financing, July 2026"
$body = "The $($P)550m financing solved a liquidity problem." + [char]10 + "The analysis says the problem is structural."
$b = Txt $s 40 120 880 62 $body 24 $true $NAVY
$bullets = @(
 "$($P)340m of pro forma liquidity bought c.19 months of runway - a window, not a solution",
 "Aston Martin does not become self-sustaining within the forecast horizon; liquidity is exhausted c. February 2028",
 "Even at a 0% cost of refinancing, FY2030 free cash flow is still $($P)(115)m - no achievable coupon fixes this structure",
 "The March 2029 maturity wall is $($P)1,357m, c.4x the liquidity the deal created",
 "The 2029 notes were marked at 79.6% of par at 30 June 2026 - the market has already priced a restructuring"
) -join [char]13
$b2 = Txt $s 40 200 880 150 $bullets 15 $false 0
$b2.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = -1
$b2.TextFrame.TextRange.ParagraphFormat.Bullet.Type = 1
$b2.TextFrame.TextRange.ParagraphFormat.Bullet.Character = 8226
CalloutBox $s 40 380 880 62 "RECOMMENDATION: The financing delayed a restructuring rather than avoiding one. Expect a further capital event - additional debt, an equity injection, an asset sale, or further shareholder support - before the March 2029 maturity." $NAVY | Out-Null
Notes $s "This slide must stand alone. Lead with the framing sentence, then the 0% refinancing test - it is the single most robust finding and does not depend on any provisional driver. Everything else on this slide is sourced from AML primary documents."

# ---------------- SLIDE 2 ----------------
$s = NewSlide "Why the financing was needed - and what was done" "An urgent liquidity problem, met at high cost and at speed"
Txt $s 40 118 420 24 "POSITION AT 30 JUNE 2026" 13 $true $NAVY | Out-Null
$d = @(
 @("Metric","Reported"),
 @("Net debt","$($P)1,544.7m"),
 @("Total liquidity","$($P)145.2m"),
 @("Adjusted net leverage","8.9x"),
 @("H1 2026 adjusted EBIT","$($P)(108.9)m"),
 @("H1 2026 free cash outflow","$($P)(197.6)m"),
 @("RCF drawn","$($P)163.0m of $($P)170m - fully utilised")
)
Tbl $s 7 2 40 146 420 224 $d @(230,190) | Out-Null
Txt $s 500 118 420 24 "THE TRANSACTION - 22 JULY 2026" 13 $true $NAVY | Out-Null
$terms = @(
 "$($P)450m senior secured term loan plus a $($P)100m delayed-draw facility",
 "Priced at SONIA + 6.75%, maturing July 2031",
 "Lead lender: HPS Investment Partners; a further $($P)100m of junior debt capacity permitted",
 "Repaid the $($P)170m RCF and the $($P)20m drawn Yew Tree facility - both cancelled",
 "Pro forma liquidity raised to c.$($P)340m"
) -join [char]13
$t2 = Txt $s 500 146 420 200 $terms 13 $false 0
$t2.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = -1
$t2.TextFrame.TextRange.ParagraphFormat.Bullet.Type = 1
$t2.TextFrame.TextRange.ParagraphFormat.Bullet.Character = 8226
CalloutBox $s 40 392 880 54 "Guidance for FY2026 net cash interest was raised to c.$($P)160m from c.$($P)150m on the back of the transaction. The full-year run-rate cost is higher still: FY2027 cash interest of $($P)191m, an increase of $($P)31m, because FY2026 carries only c.5 months of the new loan." $NAVY | Out-Null
Notes $s "All figures on this slide are sourced from the H1 2026 interim results. The RCF nuance is worth knowing: $($P)163.0m was drawn in cash and a further $($P)5.9m reserved for letters of credit, which is why AML describes it as fully utilised. Both statements are correct."

# ---------------- SLIDE 3 ----------------
$s = NewSlide "What the deal changed - and what it did not" "Liquidity up $($P)197m. The 2029 maturity wall untouched."
$d = @(
 @("Nominal, $($P)m","Pre 30-Jun-26","Post pro forma"),
 @("USD SSN 10.0% (Mar-29)","792.3","792.3"),
 @("GBP SSN 10.375% (Mar-29)","565.0","565.0"),
 @("New senior secured term loan (Jul-31)","-","450.0"),
 @("RCF and Yew Tree facility","183.0","- (cancelled)"),
 @("Other debt and leases","135.3","135.3"),
 @("Gross debt","1,675.6","1,942.6"),
 @("Total liquidity","145.2","341.9")
)
$tt = Tbl $s 8 3 40 120 500 250 $d @(250,125,125)
$tt.Table.Rows(7).Cells(1).Shape.TextFrame.TextRange.Font.Bold = -1
$tt.Table.Rows(8).Cells(1).Shape.TextFrame.TextRange.Font.Bold = -1
Txt $s 570 120 350 24 "MATURITY PROFILE ($($P)m)" 13 $true $NAVY | Out-Null
$mat = @(@("Year","Amount"),@("2026","46"),@("2027","-"),@("2028","-"),@("2029","1,357"),@("2030","-"),@("2031","450"))
$mt = Tbl $s 7 2 570 146 350 224 $mat @(175,175)
$mt.Table.Rows(5).Cells(1).Shape.TextFrame.TextRange.Font.Bold = -1
$mt.Table.Rows(5).Cells(2).Shape.TextFrame.TextRange.Font.Bold = -1
$mt.Table.Rows(5).Cells(2).Shape.TextFrame.TextRange.Font.Color.RGB = $RED
$imp = "IMPROVED: the 2028 RCF maturity is removed, $($P)450m is pushed out to 2031, and liquidity rises $($P)197m." + [char]13 + "NOT ADDRESSED: the March 2029 wall is unchanged, gross debt rises $($P)267m, and cash interest rises $($P)31m."
Txt $s 40 388 880 60 $imp 14 $true 0 | Out-Null
Notes $s "The maturity table is the exhibit - if this were rebuilt as a bar chart the 2029 bar would dominate the page, which is the point. The deal genuinely improved the maturity profile at the short end. It did nothing about the wall."

# ---------------- SLIDE 4 ----------------
$s = NewSlide "The structural test" "This is not a refinancing problem. It is a business problem."
Txt $s 40 116 430 22 "TEST 1 - REFINANCING THE 2029 NOTES ACROSS THE RATE CURVE (FY2030)" 11 $true $NAVY | Out-Null
$d = @(
 @("Refinancing rate","Coverage","FY30 free cash flow"),
 @("0.00% (hypothetical)","5.89x","$($P)(115.1)m"),
 @("10.375% (existing coupon)","1.61x","$($P)(255.9)m"),
 @("20.16% (market-implied)","0.96x","$($P)(388.7)m")
)
$t1 = Tbl $s 4 3 40 140 430 110 $d @(160,90,180)
$t1.Table.Rows(2).Cells(1).Shape.TextFrame.TextRange.Font.Bold = -1
$t1.Table.Rows(2).Cells(3).Shape.TextFrame.TextRange.Font.Bold = -1
$t1.Table.Rows(2).Cells(3).Shape.TextFrame.TextRange.Font.Color.RGB = $RED
Txt $s 40 258 430 40 "Even at a zero-percent coupon, free cash flow is still $($P)(115)m. Refinancing terms are not the binding constraint." 13 $true $RED | Out-Null
Txt $s 500 116 420 22 "TEST 2 - CAN THE BUSINESS FUND ITS OWN CAPEX?" 11 $true $NAVY | Out-Null
$d2 = @(
 @("$($P)m","FY27","FY28","FY29","FY30","FY31"),
 @("EBITDA","288","300","306","313","320"),
 @("Capex","(350)","(350)","(350)","(350)","(300)"),
 @("EBITDA less capex","(62)","(51)","(44)","(37)","20")
)
$t2 = Tbl $s 4 6 500 140 420 110 $d2 @(120,60,60,60,60,60)
$t2.Table.Rows(4).Cells(1).Shape.TextFrame.TextRange.Font.Bold = -1
Txt $s 500 258 420 40 "EBITDA does not cover capex alone until FY2031 - and then by $($P)20m, before a penny of interest." 13 $true $RED | Out-Null
Txt $s 40 312 880 44 "TEST 3 - THE MARKET'S OWN VERDICT: the 2029 notes were marked at 79.6% of par at 30 June 2026, down from 92.6% in December and before this financing was announced. An implied yield of c.20% prices a restructuring, not a refinancing." 13 $false 0 | Out-Null
CalloutBox $s 40 368 880 76 "CAVEAT: 20.16% is the market-implied DISTRESS yield on the 2029 notes at 30 June 2026. It embeds default probability and is not a forecast new-issue coupon. The conclusion does not depend on it - the 0% row makes the same point. All three tests above are independent of the forecast assumptions." $RED | Out-Null
Notes $s "This is the load-bearing slide. Every finding here is assumption-light: the 0% refinancing row, the EBITDA-less-capex line and the observed bond price all hold regardless of the refinancing rate, the SONIA path or the covenant threshold. If challenged on the forecast drivers, retreat to this slide - it survives."

# ---------------- SLIDE 5 ----------------
$s = NewSlide "Scenario output: no margin for error" "The base case is already cash-negative. The downside only changes the date."
$d = @(
 @("","Base case","Downside case"),
 @("Covenant breach","FY2027","FY2027"),
 @("Liquidity exhausted","c. February 2028","c. November 2027"),
 @("Runway from 22-Jul-2026","c.19 months","c.16 months"),
 @("Peak funding gap","$($P)1,319m","$($P)2,163m"),
 @("FY2031 net debt / EBITDA","10.2x","29.9x"),
 @("FY2029 interest coverage","1.0x","0.4x")
)
$t = Tbl $s 7 3 40 118 540 224 $d @(230,155,155)
$t.Table.Rows(3).Cells(1).Shape.TextFrame.TextRange.Font.Bold = -1
$t.Table.Rows(3).Cells(2).Shape.TextFrame.TextRange.Font.Bold = -1
$t.Table.Rows(3).Cells(3).Shape.TextFrame.TextRange.Font.Bold = -1
$rf = @(
 "The base case is not a comfortable case: it assumes recovery to Aston Martin's own best-ever EBITDA margin (17.1%, FY2024) and still burns cash in every forecast year.",
 "On a volume-only stress the model tips into liquidity exhaustion at a decline of just 6.6% - the base case sits on the cliff edge.",
 "Covenant breach occurs in FY2027 at every tested threshold ($($P)75m / $($P)100m / $($P)125m), so the date does not depend on the undisclosed covenant level."
) -join [char]13
$rb = Txt $s 610 118 310 220 $rf 12 $false 0
$rb.TextFrame.TextRange.ParagraphFormat.Bullet.Visible = -1
$rb.TextFrame.TextRange.ParagraphFormat.Bullet.Type = 1
$rb.TextFrame.TextRange.ParagraphFormat.Bullet.Character = 8226
CalloutBox $s 40 360 880 76 "CAVEAT: the downside is a NO-RESPONSE case. It deliberately excludes the mitigating actions management states it would take - capex deferral first - which matches the construction of management's own disclosed reverse stress test. Real mitigation would soften these outcomes. The base case, however, requires no stress at all to fail." $RED | Out-Null
Notes $s "The reframing is the point of this slide. The story is not 'the downside is severe' - it is 'the base case has no margin for error'. Note also that the base case is deliberately more conservative than management's undisclosed plan; that divergence is quantified in the supporting analysis and should be conceded openly if challenged."

# ---------------- SLIDE 6 ----------------
$s = NewSlide "Alternatives considered - and recommendation" "The right deal available. The wrong problem solved."
$d = @(
 @("Option","Assessment"),
 @("Rights issue","Only option that fixes solvency - but c.$($P)1.3bn needed, roughly equal to FY2025 revenue. Shareholder base already stretched"),
 @("Convertible bond","Saves cash interest; does not touch the $($P)62m FY2027 EBITDA-less-capex deficit. Defers, does not fix"),
 @("Public secured bond","Effectively unplaceable - the existing notes were trading at 79.6% of par"),
 @("Private credit (actual)","Fast, certain, flexible. Cost: $($P)31m more interest a year, more secured claims, c.19 months bought"),
 @("Mixed debt / equity","The only structure addressing both the liquidity gap and the structural deficit")
)
Tbl $s 6 2 40 116 880 168 $d @(180,700) $true | Out-Null
CalloutBox $s 40 292 430 66 "MARKET CONTEXT (per Bloomberg reporting; not confirmed by Aston Martin): the financing was reported as a drop-down structure subordinating existing bondholders, with a creditor group led by Arini, BlackRock and Sculptor opposing it and retaining Jefferies. Illustrates why the alternatives were harder to execute. Not relied upon in any figure or conclusion." $GREY | Out-Null
$rec = "RECOMMENDATION - the financing delayed a restructuring rather than avoiding one." + [char]13 + "As execution the deal is defensible: with the public market shut and shareholders stretched, private credit was plausibly the only executable option. As a solution it does not work - it addressed liquidity; the deficit is structural." + [char]13 + "Expect a further capital event - additional debt, an equity injection, an asset sale, or further shareholder support - before the March 2029 maturity."
Txt $s 500 292 420 130 $rec 12 $true $NAVY | Out-Null
Txt $s 40 372 430 60 "What would have to be true for the softer verdict: FY2031 EBITDA of $($P)652m against $($P)320m modelled - c.2.4x the FY2024 peak of $($P)271m. A bet on operational transformation, not on the financing." 11 $false $GREY | Out-Null
Notes $s "Separate the two judgements explicitly - execution quality versus problem solved. That distinction is what makes the recommendation defensible rather than merely negative. The 'what would have to be true' line makes the call falsifiable: name the threshold, and the conclusion can be tested rather than argued."

$pres.SaveAs($out)
$pres.Close()
$pp.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($pp) | Out-Null
"SAVED: $out"
