$ErrorActionPreference='Stop'
$P = [char]0x00A3
$out = "C:\Users\harry\aston-martin-financing-analysis\AML_Recommendation_Memo.docx"
if (Test-Path $out) { Remove-Item $out -Force }

$NAVY = 4401680
$RED  = 192
$GREY = 6579300

$wd = New-Object -ComObject Word.Application
$wd.Visible = $false
$doc = $wd.Documents.Add()
$doc.PageSetup.TopMargin = [single]38
$doc.PageSetup.BottomMargin = [single]38
$doc.PageSetup.LeftMargin = [single]48
$doc.PageSetup.RightMargin = [single]48
$sel = $wd.Selection
$sel.Font.Name = "Calibri"

function Para($text, $size, $bold, $color, $spaceAfter, $align) {
  $sel.Font.Size = [single]$size
  if ($bold) { $sel.Font.Bold = [int]1 } else { $sel.Font.Bold = [int]0 }
  $sel.Font.Italic = [int]0
  if ($color -ne $null) { $sel.Font.Color = [int]$color } else { $sel.Font.Color = [int]0 }
  $sel.ParagraphFormat.SpaceAfter = [single]$spaceAfter
  $sel.ParagraphFormat.SpaceBefore = [single]0
  if ($align -ne $null) { $sel.ParagraphFormat.Alignment = [int]$align } else { $sel.ParagraphFormat.Alignment = [int]3 }
  $sel.TypeText([string]$text)
  $sel.TypeParagraph()
}
function Head($text) {
  $sel.Font.Size = [single]11
  $sel.Font.Bold = [int]1
  $sel.Font.Color = [int]$NAVY
  $sel.Font.Italic = [int]0
  $sel.ParagraphFormat.SpaceAfter = [single]2
  $sel.ParagraphFormat.SpaceBefore = [single]5
  $sel.ParagraphFormat.Alignment = [int]0
  $sel.TypeText([string]$text.ToUpper())
  $sel.TypeParagraph()
}
function Rule() {
  $sel.ParagraphFormat.SpaceAfter = 6
  $sel.ParagraphFormat.SpaceBefore = 2
  $sel.Font.Color = [int]0
  $sel.TypeText(" ")
  $sel.Paragraphs.Item(1).Range.ParagraphFormat.Borders.Item(-3).LineStyle = 1
  $sel.TypeParagraph()
}

# Header
Para "INVESTMENT COMMITTEE MEMORANDUM" 13 $true $NAVY 2 0
Para "Aston Martin Lagonda Global Holdings plc" 11 $true 0 4 0

$sel.Font.Size = [single]9.5
$sel.Font.Bold = [int]0
$sel.Font.Color = [int]$GREY
$sel.ParagraphFormat.SpaceAfter = [single]1
$sel.ParagraphFormat.Alignment = [int]0
$sel.TypeText("TO: Investment Committee      FROM: Credit Analysis      DATE: August 2026")
$sel.TypeParagraph()
$sel.TypeText("RE: $($P)550m private credit financing (July 2026) - solved, or postponed?")
$sel.TypeParagraph()
$sel.Font.Color = [int]0

Head "Recommendation"
Para "The July 2026 financing delayed a restructuring rather than avoiding one. It solved a liquidity problem; the analysis indicates the problem is structural. We expect a further capital event - additional debt, an equity injection, an asset sale, or further shareholder support - before the March 2029 maturity." 9.8 $true 0 3 3

Head "Situation"
Para "At 30 June 2026 Aston Martin carried $($P)1,544.7m of net debt against $($P)145.2m of total liquidity, adjusted net leverage of 8.9x, an H1 adjusted EBIT loss of $($P)108.9m and an H1 free cash outflow of $($P)197.6m. The $($P)170m revolving credit facility was fully utilised. On 22 July the Group closed a $($P)450m senior secured term loan plus a $($P)100m delayed-draw facility at SONIA + 6.75%, maturing July 2031, led by HPS Investment Partners, repaying and cancelling both the RCF and the $($P)20m drawn Yew Tree facility. Pro forma liquidity rose to c.$($P)340m." 9.8 $false 0 3 3

Head "Analysis"
Para "The financing does not change the trajectory. On a base case capped at Aston Martin's own best-ever EBITDA margin (17.1%, FY2024) and carrying its own guided $($P)1.7bn capital programme, free cash flow is negative in every forecast year. Liquidity is exhausted around February 2028 - a runway of roughly 19 months - with a covenant breach in FY2027 at every plausible threshold." 9.8 $false 0 3 3
Para "Refinancing terms are not the binding constraint. Testing the 2029 maturity across the rate curve, even a hypothetical 0% coupon leaves FY2030 free cash flow at $($P)(115)m. The company does not cover capital expenditure alone until FY2031, and then by only $($P)20m - before any interest. The 2029 wall of $($P)1,357m is c.4x the liquidity the deal created." 9.8 $false 0 3 3
Para "The market has already priced this. The 2029 notes were marked at 79.6% of par at 30 June 2026, down from 92.6% in December and before this financing was announced - an implied yield of c.20%, which prices a restructuring rather than a refinancing. (This is a distress yield reflecting default risk, not a forecast new-issue coupon.)" 9.8 $false 0 3 3
Para "There is no margin for error. On a volume-only stress the base case tips into liquidity exhaustion at a decline of just 6.6%. A combined downside - anchored on management's own disclosed reverse stress test - brings exhaustion forward to c.November 2027 and lifts the peak funding requirement from $($P)1.3bn to $($P)2.2bn. (The downside excludes the mitigating actions management states it would take; real mitigation would soften it.)" 9.8 $false 0 3 3

$sel.Font.Size = [single]9
$sel.Font.Italic = [int]1
$sel.Font.Color = [int]$GREY
$sel.ParagraphFormat.SpaceAfter = [single]4
$sel.ParagraphFormat.LeftIndent = [single]14
$sel.TypeText("Market context (per Bloomberg reporting; not confirmed by Aston Martin): the financing was reported as a drop-down structure subordinating existing bondholders, with a creditor group led by Arini, BlackRock and Sculptor opposing it and retaining Jefferies. This illustrates why the alternatives were harder to execute. It is not relied upon in any figure or conclusion above.")
$sel.TypeParagraph()
$sel.ParagraphFormat.LeftIndent = [single]0
$sel.Font.Italic = [int]0
$sel.Font.Color = [int]0

Head "Assessment of the decision"
Para "Two judgements should be separated. As execution, the transaction is defensible: with the public market effectively shut, a shareholder base already stretched and an urgent need, private credit was plausibly the only option available in the window. As a solution, it does not work - it bought c.19 months at a cost of $($P)31m of additional annual interest, further secured claims ranking ahead of existing bondholders, and a materially more encumbered balance sheet." 9.8 $false 0 3 3

Head "Key risks to this view"
Para "The forecast beyond FY2026 rests on drivers Aston Martin has not guided; our base case is deliberately more conservative than management's undisclosed plan, and the conclusion is sensitive to it. For the softer 'bought sufficient time' verdict to hold, FY2031 EBITDA would need to reach $($P)652m - more than double our modelled $($P)320m, and c.2.4x the company's FY2024 peak of $($P)271m." 9.8 $false 0 3 3

$sel.Font.Size = [single]8
$sel.Font.Italic = [int]1
$sel.Font.Color = [int]$GREY
$sel.ParagraphFormat.SpaceBefore = [single]5
$sel.TypeText("Prepared from Aston Martin Lagonda primary disclosures (H1 2026 interim results, Q1 2026 results, Annual Report 2025) and Bank of England published rates. Full sourcing, derivations and caveats are set out in the accompanying model and sourcing record.")

$doc.SaveAs([ref]$out, [ref]16)
$pages = $doc.ComputeStatistics(2)
"pages: $pages"
$doc.Close()
$wd.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wd) | Out-Null
"SAVED: $out"
