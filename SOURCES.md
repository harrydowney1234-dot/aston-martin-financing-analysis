# SOURCES — verified figures with primary-source quotes

Every figure that will enter the model, tied to an Aston Martin primary document.
Rule: figures are read from locally extracted PDF text, never from a web summariser.

## Documents used

> **Note:** the source PDFs are **not committed to this repository** — they are Aston Martin's
> copyrighted documents and are freely available from
> [Aston Martin investor relations](https://www.astonmartin.com/en/corporate/investors). Every figure
> below carries a verbatim quote, so the audit trail is complete without them. The `Local file` column
> records the working-copy filename used during the analysis.

| Ref | Document | Local file | Origin |
| --- | --- | --- | --- |
| **[H1]** | AML Interim results, six months ended 30 June 2026 | `source_docs/aston-martin-lagonda-h1-2026-results.pdf` | astonmartin.com |
| **[H1P]** | H1 2026 results presentation | `source_docs/aston-martin-lagonda-h1-2026-results-presentation.pdf` | astonmartin.com |
| **[H1T]** | H1 2026 results transcript | `source_docs/aston-martin-lagonda-h1-2026-results-transcript.pdf` | astonmartin.com |
| **[Q1]** | AML Q1 2026 results, three months ended 31 March 2026 | `source_docs/aston-martin-lagonda-q1-2026-results.pdf` | astonmartin.com |
| **[AR25]** | AML Annual Report 2025 | `source_docs/aston-martin-lagonda-annual-report-2025.pdf` | astonmartin.com |
| **[PRESS]** | Bloomberg articles — **deal context only, never a model input** | n/a | bloomberg.com |

---

## 1. Performance — H1 2026

| Metric | Value | Comparative | Ref |
| --- | --- | --- | --- |
| Total wholesale volumes | 2,331 (+21%) | H1 2025: 1,922 | [H1] |
| Revenue | £628.6m (+38%) | H1 2025: £454.4m | [H1] |
| Gross profit | £212.5m (+68%) | H1 2025: £126.6m | [H1] |
| Gross margin | 33.8% | H1 2025: 27.9% | [H1] |
| Adjusted EBITDA | £62.7m | H1 2025: £(3.0)m | [H1] |
| Adjusted EBITDA margin | 10.0% | H1 2025: (0.7)% | [H1] |
| **Adjusted EBIT** | **£(108.9)m** | H1 2025: £(121.5)m | [H1] |
| Operating loss | £(56.5)m | H1 2025: £(134.7)m | [H1] |
| Loss before tax | £(154.2)m | H1 2025: £(140.8)m | [H1] |
| Adjusted opex (ex D&A) | £150m (+16%) | H1 2025: £130m | [H1] |

> "Adjusted EBITDA increased by £66m in H1 2026 to £63m (H1 2025: £(3)m)… adjusted EBITDA margin
> increased to 10% (H1 2025: (1)%). Adjusted EBIT improved by 10% in H1 2026 to £(109)m" — [H1]

## 2. Cash flow — H1 2026

| Metric | H1 2026 | H1 2025 | Q2 2026 | Ref |
| --- | --- | --- | --- | --- |
| Cash (used in)/from operating activities | £(2.3)m | £(81.0)m | £50.9m | [H1] |
| Cash used in investing activities | £(120.2)m | £(170.6)m | £(59.2)m | [H1] |
| Net cash interest paid | £(75.1)m | £(69.4)m | £(72.5)m | [H1] |
| **Free cash outflow** | **£(197.6)m** | £(321.0)m | £(80.8)m | [H1] |
| Q1 2026 free cash outflow | £(116.8)m | Q1 2025: £(120.3)m | — | [Q1] |

> "Free cash outflow in H1 2026 of £198m materially improved compared to the prior year period
> (H1 2025: £321m outflow)" — [H1]

**Check:** Q1 £116.8m + Q2 £80.8m = £197.6m = H1. ✅ ties

## 3. Capital structure @ 30 June 2026

### Net debt bridge

| Component | 30 Jun 26 | 31 Dec 25 | 30 Jun 25 | Ref |
| --- | --- | --- | --- | --- |
| Loan notes | (1,345.2) | (1,329.8) | (1,310.6) | [H1] |
| Inventory financing | (39.1) | (39.6) | (38.0) | [H1] |
| Bank loans and overdrafts | (168.8) | (170.4) | (58.7) | [H1] |
| Committed facility (YTC) | (18.3) | – | – | [H1] |
| Lease liabilities (IFRS 16) | (89.7) | (91.8) | (94.0) | [H1] |
| **Gross debt** | **(1,661.1)** | (1,631.6) | (1,501.3) | [H1] |
| Cash balance | 114.9 | 249.9 | 123.6 | [H1] |
| Cash not available for short-term use | 1.5 | 1.4 | – | [H1] |
| **Net debt** | **(1,544.7)** | (1,380.3) | (1,377.7) | [H1] |

**Check:** 1,661.1 − 114.9 − 1.5 = 1,544.7 ✅ ties

### Loan notes by tranche @ 30 June 2026

| Instrument | Nominal £m | Book £m | Fair value £m | FV % of nominal |
| --- | --- | --- | --- | --- |
| $1,050.0m 10% USD Notes | 792.3 | 786.3 | 629.9 | 79.5% |
| £465.0m 10.375% GBP Notes | 465.0 | 461.6 | 372.3 | 80.1% |
| £100.0m 10.375% GBP Notes | 100.0 | 97.3 | 77.8 | 77.8% |
| **Total** | **1,357.3** | **1,345.2** | **1,080.0** | **79.6%** |

**Check:** book value total 1,345.2 = "Loan notes" line in net debt bridge ✅ ties
**Implied FX @ 30 Jun 2026:** $1,050.0m ÷ £792.3m = **1.3253 USD/GBP** (derived, not stated)

> "$1,050.0m Senior Secured Notes ("SSNs") at 10.0% and £565.0m of SSNs at 10.375% both of which
> mature in March 2029" — [H1], Going Concern

**Note:** the £565m is **two tranches** (£465m + £100m), both 10.375%, both March 2029.

**Market signal:** at 31 Dec 2025 the same notes were marked at £1,243.5m vs £1,343.3m nominal (92.6%).
By 30 Jun 2026 — *before* the 22 July financing — they had fallen to **79.6% of par**.

### Facilities

| Facility | Size | Drawn @ 30 Jun 26 | Status | Ref |
| --- | --- | --- | --- | --- |
| Super senior RCF | £170.0m | £163.0m cash + £5.9m reserved for LCs/guarantees | Repaid in full and cancelled 22 Jul 2026 | [H1] |
| YTC Committed Facility | £50.0m | £20.0m (£18.3m net of £1.7m unamortised fees) | Repaid in full and cancelled 22 Jul 2026 | [H1] |
| Inventory repurchase | — | £39.1m (incl. £1.1m accrued interest) | **Repayable October 2026, £40.0m gross** | [H1] |

> "At 30 June 2026 £163.0m of the £170.0m RCF was drawn down in cash… £5.9m of the RCF has been
> reserved for the issuance of letters of credit and guarantees" — [H1]

### Liquidity and leverage

| Metric | 30 Jun 26 | 31 Dec 25 | 30 Jun 25 | Ref |
| --- | --- | --- | --- | --- |
| Cash balance | 114.9 | 249.9 | 123.6 | [H1] |
| Available facilities | 30.3 | — | 104.1 | [H1] |
| **Total liquidity** | **145.2** | 250 | 227.7 | [H1] |
| Adjusted LTM EBITDA | 173.8 | 108.1 | 205.8 | [H1] |
| **Adjusted net leverage** | **8.9x** | 12.8x | 6.7x | [H1] |

**Check:** 1,544.7 ÷ 173.8 = 8.89x ✅ · 1,380.3 ÷ 108.1 = 12.77x ✅ · 1,377.7 ÷ 205.8 = 6.69x ✅

**Definitions** (verbatim, [H1] Appendix):
> "Net Debt is current and non-current borrowings in addition to inventory financing arrangements,
> lease liabilities, less cash and cash equivalents and cash held not available for short-term use"
> "Adjusted net leverage is represented by the ratio of Net Debt to the last twelve months ('LTM')
> Adjusted EBITDA"
> "Free cash flow is represented by cash (outflow)/inflow from operating activities plus the cash used
> in investing activities (excluding interest received, proceeds from disposal of investments and gross
> proceeds from the disposal of internally generated assets less cash settled fees in the period) plus
> interest paid in the period less interest received"

## 4. The July 2026 financing — all terms confirmed from primary source

> "On 22 July 2026, the Group announced the closing of a new £550m debt financing (the "Financing")…
> The Financing consists of a £450m Senior Secured Term Loan ("SSTL") and a £100m Delayed Draw Term
> Loan, priced at 6.75% over the prevailing SONIA base rate and maturing July 2031, with lead lenders
> being investment funds and accounts managed by HPS Investment Partners ("HPS"). There is an
> additional £100m permitted debt incurrence capacity, junior to the Financing. The £450m gross
> proceeds from the SSTL have been used to repay both the Group's fully utilised £170m super senior
> revolving credit facility ("RCF") and the £20m drawn under the £50m facility committed by members of
> the Yew Tree Consortium ("YTC Facility") and to pay transaction costs, with the balance for general
> corporate purposes. The existing RCF commitments and the YTC Facility were simultaneously cancelled.
> The SSTL enhances the Group's pro forma liquidity as at 30 June 2026 to c. £340m" — [H1]

| Term | Value | Status |
| --- | --- | --- |
| SSTL | £450m | ✅ primary |
| Delayed Draw Term Loan | £100m | ✅ primary |
| Pricing | SONIA + 6.75% | ✅ primary |
| Maturity | July 2031 | ✅ primary |
| Lead lender | HPS Investment Partners | ✅ primary |
| Additional junior debt capacity | £100m | ✅ primary |
| Pro forma liquidity @ 30 Jun 26 | c.£340m | ✅ primary |

## 5. FY2026 guidance ([H1], unchanged from earlier guidance)

| Item | Guidance | FY2025 actual |
| --- | --- | --- |
| Adjusted opex (ex D&A) | below £300m | £262m |
| Adjusted D&A | £375m–£400m | £297m |
| Adjusted EBIT margin | materially improve, towards breakeven | (15.0)% |
| **Net cash interest** | **c.£160m** (previously c.£150m) | — |
| Capital investment | c.£300m | £341m |
| Capex programme FY26–FY30 | c.£1.7bn (previously c.£2bn) | — |
| Free cash outflow | materially improve | £410m outflow |
| Valhalla deliveries | c.500 | — |

> "Net cash interest is expected to be c. £160m3 (previously c. £150m3)" — [H1]
> footnote 3: "Net cash interest assuming current exchange rates prevail for FY 2026"

## 6. FY2025 base year ([AR25])

| Metric | FY 2025 | FY 2024 |
| --- | --- | --- |
| Total wholesale volumes | 5,448 (−10%) | 6,030 |
| Revenue | £1,257.7m (−21%) | £1,583.9m |
| Gross profit | £369.8m | £583.9m |
| Gross margin | 29.4% | 36.9% |
| Adjusted EBITDA | £108.1m | £271.0m |
| Adjusted EBITDA margin | 8.6% | 17.1% |
| Adjusted EBIT | £(189.2)m | £(82.8)m |
| Operating loss | £(259.2)m | £(99.5)m |
| Loss before tax | £(363.9)m | £(289.1)m |
| Year-end total liquidity | £250m | — |

## 7. Q1 2026 ([Q1])

| Metric | Q1 2026 | Q1 2025 |
| --- | --- | --- |
| Total wholesale volumes | 939 (−1%) | 950 |
| Revenue | £270.4m (+16%) | £233.9m |
| Gross margin | 34.7% | 27.9% |
| Adjusted EBIT | £(56.9)m | £(64.5)m |
| Net debt | £(1,459.2)m | £(1,267.4)m |
| Total liquidity | £177.7m | £387.2m |
| Free cash outflow | £(116.8)m | £(120.3)m |

Also in Q1: £50m committed facility agreed with Yew Tree Consortium members (Yew Tree Overseas Limited
and Saint Alexander S.à r.l); £50m cash received for sale of Aston Martin F1 naming rights to AMR GP
(completed 9 March 2026, recognised as a £50.0m gain less £2.3m fees).

## 8. Covenants / going concern ([H1])

**Resolved from primary source — the RNS was not required.**

> "Under the Senior Secured Term Loan, the Group will be required to comply with a **minimum liquidity
> covenant tested monthly from August 2026**, whereby the Group must hold certain minimum levels of
> liquidity. The Group expects to be compliant with covenant requirements for the remainder of the
> going concern review period through to 30 September 2027." — [H1], Going Concern

> "Both of these commitments have now been cancelled post the period end and therefore **the Group is
> no longer required to test the RCF leverage covenant**." — [H1]

| Covenant | Status |
| --- | --- |
| SSTL minimum liquidity | Tested **monthly from August 2026**. ⚠️ **Quantum not disclosed** — see limitation below |
| RCF leverage covenant | No longer tested (RCF cancelled 22 Jul 2026) |
| Going concern review period | Through to 30 September 2027 |

### ⚠️ LIMITATION — covenant quantum

The **minimum liquidity level itself is not disclosed** in any AML primary document. Attempts made
(then stopped per instruction, rather than searching indefinitely):
- `/en/corporate/investors/regulatory-news` — the announcement list renders client-side; not retrievable
- `/en/corporate/investors/funding` — no 2026 entry (latest is November 2024)
- H1 results, presentation and transcript — covenant *type* and *frequency* disclosed, quantum not

The model will therefore show the liquidity path against a **user-selectable minimum liquidity
threshold**, not a sourced one. This is a disclosed limitation, not an assumption presented as fact.

### Management's own reverse stress test — anchor for the downside case

> "during the going concern period total core vehicle volumes (DBX and GT/Sports) would need to reduce
> by **more than 55% from forecast levels to result in having no liquidity**, and **more than 25% to
> result in a breach of covenants**." — [H1]

This is management's own quantified stress sensitivity and is the natural calibration anchor for the
Phase 3 downside case. To be proposed for sign-off, not adopted silently.

---

## 8a. Flagged assumptions — signed off by user 2026-08-07, unverified

### A1. Transaction costs in the pro forma liquidity bridge

The pro forma liquidity bridge does not tie to the stated c.£340m:

| Line | £m |
| --- | --- |
| Cash @ 30 Jun 2026 | 114.9 |
| SSTL gross proceeds | +450.0 |
| Repay RCF drawn position | (163.0) |
| Repay YTC drawn position | (20.0) |
| **Subtotal** | **381.9** |
| Implied transaction costs / other | **(c.42)** |
| **Stated pro forma liquidity** | **c.340** |

AML discloses only "and to pay transaction costs" — **no figure is given**. Note the repayment is
disclosed both as the "fully utilised £170m RCF" and, in the Going Concern note, as "the RCF drawn
position of £163.0m"; using £163.0m (the cash position actually repaid) gives the £381.9m subtotal above.

**Status: user-approved flagged assumption.** Unverified pending the FY2026 Annual Report financing
cash flows. Modelled as a sensitivity band, **not** a point estimate:

| Case | Transaction costs | Pro forma liquidity |
| --- | --- | --- |
| Low | £30m | c.£352m |
| **Central** | **£40m** | **c.£342m** |
| High | £50m | c.£332m |

The £35m originally derived sits inside this band. The central £40m is the closest fit to the
disclosed c.£340m.

### A2. The £100m DDTL is EXCLUDED from pro forma liquidity — ✅ now primary-sourced

Confirmed directly by the CFO, so this is no longer an inference:

> **Doug Lafferty (CFO):** "pro forma liquidity, therefore at the end of June standing at £340m…
> Just to reiterate, **the delayed draw term loan element of that is not included in that liquidity
> number**." — [H1T], H1 2026 results call

**Treatment: memo item only.** The £100m DDTL is committed-but-undrawn capacity sitting *on top* of
the c.£340m headline. It must **not** be added into headline liquidity in any output. It carries its
own sensitivity (drawn / undrawn) and its own footnote wherever liquidity is presented.

⚠️ Drawing it would add £100m of debt at SONIA + 6.75% — liquidity relief, not deleveraging. The
conditions attaching to the delayed draw are not disclosed.

## 8b. Phase 2 model inputs — external and computed (added 2026-08-07)

### B1. SONIA base path — SOURCED (Bank of England)

| Item | Value | Source |
| --- | --- | --- |
| BoE Bank Rate | **3.75%** (as at July 2026) | [Bank of England](https://www.bankofengland.co.uk/boeapps/database/Bank-Rate.asp) |
| SONIA vs Bank Rate | SONIA sits c.**3bp below** Bank Rate | [BoE, SONIA benchmark](https://www.bankofengland.co.uk/markets/sonia-benchmark) |
| **Model input** | **3.72% held flat FY2026–FY2031** | derived from the two above |

⚠️ **LIMITATION — forward curve not obtained.** The BoE publishes an OIS/SONIA forward curve at
`bankofengland.co.uk/statistics/yield-curves`, but that page and the Bank Rate database page both
return **HTTP 403** to our user agent. A market-implied forward path could not be retrieved, so the
rate is held flat at spot rather than following a curve. This was checked, not assumed.

**Materiality is low:** SONIA applies only to the £450m SSTL (£550m if the DDTL is drawn). A ±100bp
move changes annual interest by only ±£4.5m against total cash interest of c.£190m — under 2.5%.
SONIA is therefore not a swing factor in the base case, though it is one leg of the Phase 3 downside.

### B2. Refinancing rate on the 2029 notes — COMPUTED from the sourced market price

Replaces the earlier provisional 11.0% guess. Derived from the disclosed 30 June 2026 fair values (§3).

Settlement 30-Jun-2026, maturity 31-Mar-2029 (2.75 years), semi-annual coupons, 6 payments remaining.
Dirty price = clean price + 3 months' accrued interest. Solved by bisection for y in:

> Σ (coupon/2) / (1 + y/2)^(2t) + 100 / (1 + y/2)^(2 × 2.75) = dirty price,
> for t = 0.25, 0.75, 1.25, 1.75, 2.25, 2.75

| Tranche | Clean % | Accrued | Dirty % | **Implied YTM** |
| --- | --- | --- | --- | --- |
| USD SSN 10.000% | 79.50 | 2.50 | 82.00 | **20.02%** |
| GBP SSN 10.375% (£465m) | 80.06 | 2.59 | 82.66 | **20.14%** |
| GBP SSN 10.375% (£100m) | 77.80 | 2.59 | 80.39 | **21.44%** |
| **Blended** | **79.57** | **2.54** | **82.11** | **20.16%** |

Blended coupon = (792.3 × 10.000 + 565.0 × 10.375) / 1357.3 = 10.156%.
Full workings are on the **YTM** tab of the model.

### 🔒 RULE — the caveat travels with the number (user decision, 2026-08-07)

⚠️ **20.16% is a DISTRESS yield.** It embeds the market's probability of default. It is **not** a
forecast of a clean new-issue coupon. Using it as the 2029 refinancing rate is deliberately
conservative. The honest reading is **not** "AML will pay 20%", but "the market does not currently
believe these notes can be refinanced on economic terms" — which is itself the finding.

**This caveat must appear wherever the number is quoted — model, memo, deck, or conversation.** It is
explicitly *not* sufficient for it to sit on the YTM tab alone. Standard wording:

> *"20.16% is the market-implied distress yield on the 2029 notes at 30 June 2026, not a forecast
> new-issue coupon; it embeds default probability."*

Placed in the workbook at: Assumptions I29 · Debt H38 · Outputs A4 · README B29–B30 · YTM A19–A24.
**Carry into the Phase 5 memo and deck.**

### B2a. Attribution — what actually drives the leverage/coverage reversal

Checked by re-running the model with each revised input reverted one at a time (cumulative walk).
Recorded permanently on the **Attribution** tab; step S4 reconciles exactly to the live model.

| Metric | Total swing | Margin cap | SONIA | FY26 calib. | Refi rate |
| --- | --- | --- | --- | --- | --- |
| FY2031 leverage | +4.62x | **+3.61x (78%)** | −0.02x | −0.03x | +1.06x (23%) |
| FY2029 coverage | −0.94x | −0.47x (50%) | +0.01x | 0.00x | −0.48x (51%) |
| FY2031 coverage | −1.23x | −0.64x (52%) | +0.01x | 0.00x | −0.60x (49%) |
| Peak funding gap | +£758.8m | **+£435.3m (57%)** | −£6.9m | −£11.5m | +£341.9m (45%) |

**Finding:** the *leverage* reversal is driven **mainly by the margin cap (78%)**, not the refi rate
(23%). The *coverage* collapse is roughly an even split. SONIA and the FY2026 calibration are
immaterial (under 2% each) and both slightly **reduce** stress.

This is **not** a compounding of unexamined provisional assumptions. Two deliberate corrections did the
work, and each replaced a weaker input with a better-grounded one — the margin cap is anchored to AML's
own FY2024 high (sourced), the refi rate is computed from the disclosed market price (sourced).

**Robustness:** at step S1 — margin cap applied but refi still at the old 11% — leverage already rises
to 9.17x and the funding gap reaches £995m. **The conclusion does not depend on the distress yield.**

### B3. FY2026 cash interest calibrated to guidance — explicit, not silent

The model accrues a full year of note coupons, giving £171.6m for FY2026. Company guidance is
**c.£160m** of net cash interest, reflecting actual payment dates rather than accrual. FY2026 is
calibrated down by **£11.6m** to match guidance, on a visible line (Debt tab row 42), because FY2026 is
the only year with company guidance to calibrate against.

**FY2027 onwards retains the accrual methodology** — there is no guidance to check it against, so
introducing an unanchored adjustment there would be inventing precision.

### B4. EBITDA margin cap — basis for the ceiling

AML publishes **no quantified medium-term margin target**. Searched across the H1 2026 results, the H1
results transcript and the FY2025 Annual Report for any numeric transformation-programme target: none
exists. The only forward statements are qualitative ("margin expansion", "sustainably profitable
growth").

Absent a company anchor, the base case caps the adjusted EBITDA margin at approximately **17.1% — AML's
FY2024 high** (sourced, §6: FY2024 adjusted EBITDA £271.0m on revenue £1,583.9m). The forecast peaks at
**17.3%** in FY2027 and settles at 17.0%. This is the "optimistic but historically grounded" ceiling:
the company is assumed to recover to its own recent best, and no better.

## 8c. Phase 3 downside case — derivation (built 2026-08-07, user sign-off obtained)

Single combined stress per the brief: lower deliveries, weaker margins and higher SONIA applied
**together** from FY2027, never in isolation. FY2025A and FY2026E are **not** stressed — FY25 is actual
and FY26 is covered by company guidance.

### The anchor (sourced)

> "during the going concern period total core vehicle volumes (DBX and GT/Sports) would need to reduce
> by **more than 55% from forecast levels to result in having no liquidity**, and **more than 25% to
> result in a breach of covenants**." — [H1], Going Concern

### C1. Volume stress: −23% of total wholesale volumes

Management's anchor is stated on **core volumes (DBX and GT/Sports)** — it excludes Specials. The model
runs on *total* wholesale volumes, so the anchor must be translated:

- Specials ≈ **c.500 Valhalla** deliveries guided for FY2026 [H1] out of c.5,900 total units ⇒ **c.8.5%**
- Core ≈ **91.5%** of units
- −25% on core, Specials held flat ⇒ −25% × 91.5% = **−22.9%, rounded to −23% of total volumes**

Chosen over a blunt −25% of total precisely because it is traceable back to the disclosed figure.

### C2. ASP: +5.5% vs base (mix effect only, not pricing)

Cutting core volumes while holding Specials flat raises the Specials share of a smaller book, and
Specials carry a far higher ASP. Blended ASP therefore **rises** even though nothing is assumed about
pricing. Calibrated so the combined volume-and-mix effect lands at **c.−18.8% revenue** vs base.

⚠️ This is a modelled mix consequence, not a sourced number. AML does not disclose a Specials ASP.

### C3. Gross margin: 32.0% (−350bps vs base)

Bounded by two **sourced** actuals: the FY2025 trough of **29.4%** (§6) and the H1 2026 run-rate of
**33.8%** (§1). 32.0% sits between them — a return to recent-trough profitability without inventing a
margin the company has never posted.

### C4. SONIA: 5.25% (+153bps vs base)

⚠️ **JUDGEMENT, not sourced.** No forward curve is obtainable (BoE returns 403, §8b/B1). Accepted as a
labelled judgement call on the basis of low materiality: c.**£4.5m per 100bps** on the £450m SSTL
(c.£5.5m with the DDTL drawn), against total cash interest of c.£210m in the downside.

### C5. Held flat — deliberately

Capex, adjusted opex, D&A, working capital, lease payments, cash tax and the refinancing rate are all
**unchanged from base**. Management's reverse stress test **excludes mitigating actions**, so this is
modelled as a **no-response case**. Holding opex and capex flat *is* the operating-deleverage effect.

> **For the memo:** state qualitatively that this is a no-response case and that real mitigation
> (capex deferral, cost action) would soften it materially. **Do not model it** — that would break
> traceability to the disclosed anchor.

### C6. DDTL drawn

Drawn in the downside (undrawn in base): a company under liquidity stress draws committed capacity.
Adds £100m of liquidity **and** £100m of debt at SONIA + 6.75%.

⚠️ Consequence: the Debt tab's "CHECK vs disclosed c.£340m" is calibrated to the **base** case, which is
how AML disclosed it. It will not tie in the downside, by construction.

### Result

| Metric | Base | Downside |
| --- | --- | --- |
| FY2027 revenue | £1,671.4m | £1,357.8m (−18.8%) |
| FY2027 adjusted EBITDA | £288.3m | £129.5m (−55%) |
| FY2027 EBITDA margin | 17.3% | 9.5% |
| Covenant breach | FY2027 | FY2027 |
| **Liquidity exhausted** | **FY2028** | **FY2027** |
| Peak funding gap | £1,318.8m | £2,162.5m |
| FY2031 net debt / EBITDA | 10.2x | 29.9x |
| FY2029 EBITDA / cash interest | 1.0x | 0.4x |

## 8d. DEFEND-LIST — why the model exhausts liquidity at a milder shock than management's stress test

**The challenge:** our downside uses −23% total volumes (management's −25% *core*) yet fully exhausts
liquidity, while management's disclosure implies **−55% core** is needed to exhaust liquidity.
Roughly double our stress. Full workings on the **Reconciliation** tab.

### Test 1 — lever walk: it is NOT the margin assumption

FY2027 year-end liquidity, each lever added cumulatively:

| Variant | FY27 EBITDA | FY27 liquidity | Change | Exhausted |
| --- | --- | --- | --- | --- |
| Base (no stress) | £288.3m | **£29.7m** | – | FY2028 |
| V1 core volumes −25% **only** | £176.8m | **£(81.8)m** | −£111.5m | **FY2027** |
| V2 + gross margin 32% | £129.3m | £(129.3)m | −£47.5m | FY2027 |
| V3 + SONIA 5.25% | £129.3m | £(136.2)m | −£6.9m | FY2027 |
| V4 + DDTL drawn (= downside) | £129.3m | £(48.2)m | +£88.0m | FY2027 |

**A volume-only −25% stress already exhausts liquidity.** The margin cut adds a further £47.5m of
damage but is *not required*. SONIA is trivial (£6.9m). Drawing the DDTL adds back £88.0m net.

### Test 2 — volume-only ladder: our model needs only −6.6%

| Core volume cut | FY27 revenue | FY27 EBITDA | FY27 liquidity | Exhausted |
| --- | --- | --- | --- | --- |
| 0% (base) | £1,671.4m | £288.3m | £29.7m | FY2028 |
| −5% | £1,608.6m | £266.0m | £7.4m | FY2028 |
| **−6.6%** | £1,588.4m | £258.9m | **£0.3m** | FY2028 ← **breakeven** |
| −10% | £1,545.7m | £243.7m | £(14.9)m | FY2027 |
| −25% (mgmt covenant marker) | £1,357.2m | £176.8m | £(81.8)m | FY2027 |
| −55% (mgmt liquidity marker) | £980.1m | £42.9m | £(215.7)m | FY2027 |

**Our model tips into FY2027 exhaustion at −6.6% core volumes, versus management's −55%.**
The gap is therefore **not in the stress levers — it is in the base case being stressed from.**

### Why: the base case is already on the cliff edge

| Base FY2027 | £m |
| --- | --- |
| Adjusted EBITDA | 288.3 |
| Cash interest | (191.0) |
| Capex | (350.0) |
| Working capital, leases, tax | (25.0) |
| **Free cash flow** | **(277.6)** |

Interest £191m + capex £350m = **£541m of largely committed outflow against £288m of EBITDA**. Base
FY2027 year-end liquidity is just £29.7m, so any material stress tips it negative.

### Test 3 — is this an artefact of our capex phasing? No

The £1.7bn FY2026–30 programme total is **sourced guidance**; only the year-by-year phasing is ours.

| Phasing FY26–30 (total £1.7bn) | FY27 liquidity | FY28 liquidity | Exhausted |
| --- | --- | --- | --- |
| A 300/350/350/350/350 (current) | £29.7m | £(236.7)m | FY2028 |
| B 300/300/350/375/375 | £79.7m | £(186.7)m | FY2028 |
| C 300/250/350/400/400 | £129.7m | £(136.7)m | FY2028 |

Re-phasing moves FY2027 by up to £100m but **cannot change the exhaustion year** — it only shifts the
burn between years.

### The four reasons the two numbers are not like-for-like

1. **Different starting forecast.** Management stresses *their* forecast; we stress *our* reconstruction.
   Ours already burns £278m in FY2027. Theirs must be materially stronger — the −6.6% vs −55% gap
   *measures* that divergence rather than contradicting it.
2. **Different horizon.** Management's going-concern window ends **30 Sep 2027**; our FY2027 ends
   **31 Dec 2027** — roughly one extra quarter of burn (c.£70–90m at our run-rate).
3. **Different metric.** Management flexes volumes only, holding per-unit economics flat. Ours also cuts
   margin and raises SONIA — though Test 1 shows volumes alone suffice.
4. **Mitigating actions.** Both exclude them, so neutral — but management states it would act, and would
   cut capex first.

### How to defend it

Do **not** claim the model contradicts management. It does not — it stresses a different, more
conservative base. The defensible statement:

> *"On a base case capped at Aston Martin's own historic best margin and carrying its own guided capex
> programme, the capital structure is already cash-negative before any stress is applied. That is the
> finding. The downside case simply confirms there is no room for error."*

⚠️ **HONEST CAVEAT:** this makes the headline conclusion sensitive to the **base case**, not just the
downside. Management's own forecast is not disclosed, so the divergence can be **bounded but not
reconciled directly**. Say so rather than implying precision that does not exist.

## 9. Deal context — PRESS ONLY, flagged, never a model input

Source: Bloomberg. Carried in the memo/deck as context, clearly attributed. Not verified against a
primary document and not used in any calculation.

- [Creditors hire Jefferies as debt worries grow](https://www.bloomberg.com/news/articles/2026-07-16/aston-martin-creditors-tap-jefferies-as-debt-worries-grow)
- [Noteholders sign cooperation pact amid debt rout](https://www.bloomberg.com/news/articles/2026-07-08/noteholders-to-aston-martin-sign-cooperation-pact-amid-debt-rout)
- [Creditor group led by Arini, BlackRock, Sculptor moves to block debt deal](https://www.bloomberg.com/news/articles/2026-07-10/aston-martin-lender-group-led-by-arini-blackrock-sculptor)
- [Creditors send legal letter over proposed debt deal](https://www.bloomberg.com/news/articles/2026-07-21/aston-martin-creditors-ramp-up-pressure-on-struggling-carmaker)
- [£550m debt deal with HPS secured](https://www.bloomberg.com/news/articles/2026-07-22/aston-martin-raises-550-million-new-debt-from-hps-to-boost-cash)
- [Creditors fume at lack of detail on debt deal](https://www.bloomberg.com/news/articles/2026-08-01/aston-martin-creditors-fume-at-lack-of-detail-on-debt-deal)

Reported (press, unverified): the financing is structured as a **drop-down**, moving assets beyond the
reach of existing creditors and pushing bondholders back in the repayment queue. A creditor group led
by Arini Capital Management, BlackRock and Sculptor Capital signed a cooperation agreement and sent a
legal letter opposing the deal.

Note the irony worth a line in the memo: BlackRock owns HPS (the new lender) and also appears in the
objecting creditor group — different funds, opposite sides of the same capital structure.

### 🔒 RULE — how this material may be used (user decision, 2026-08-07)

**Bounded role. Not load-bearing.**

| Permitted | Prohibited |
| --- | --- |
| **One** clearly labelled "market context" callout in the memo | Any input to the model |
| **One** equivalent callout in the deck | Any number, calculation or output |
| Colour for the "why private credit over the alternatives" discussion | Any part of the final recommendation |

Required label wording, or close to it:

> *"Reported bondholder friction over asset ranking, per Bloomberg reporting — not confirmed by AML."*

Every conclusion must stand on its own without this material. If removing the drop-down/bondholder
friction commentary would change any number or weaken any conclusion, the analysis is wrong and must
be rebuilt on primary-sourced facts.

**Note the defensible primary-sourced version of the same point:** the 2029 notes were marked at
**79.6% of par** at 30 June 2026 (§3). That is a disclosed, auditable market signal of refinancing
stress and needs no press attribution — it should carry the analytical weight, with the press
reporting as context only.
