# PHASE 4 — OUTPUTS

Every output the brief asks for, both scenarios. All figures from `AML_Liquidity_Model.xlsx`.
Sourcing and derivation: `SOURCES.md`. Base case = `Assumptions!B3 = 1`, downside = `2`.

⚠️ **Carry these three caveats wherever these numbers travel:**
1. The **20.16% refinancing rate is a market-implied distress yield**, not a forecast new-issue coupon
   (`SOURCES.md` §8b/B2).
2. The **downside is a no-response case** — real mitigation would soften it (`SOURCES.md` §8c/C5).
3. **FY2027–31 drivers are unsourced judgement**, capped at AML's own FY2024 best margin. The
   conclusion is sensitive to the base case, not just the stress (`SOURCES.md` §8d).

---

## 1. Revenue and EBITDA

| £m | FY2025A | FY2026E | FY2027E | FY2028E | FY2029E | FY2030E | FY2031E |
| --- | --- | --- | --- | --- | --- | --- | --- |
| **Revenue — base** | 1,257.9 | 1,598.9 | 1,671.4 | 1,745.1 | 1,792.0 | 1,839.5 | 1,887.6 |
| Revenue — downside | 1,257.9 | 1,598.9 | 1,357.8 | 1,417.7 | 1,455.7 | 1,494.3 | 1,533.4 |
| **Adj. EBITDA — base** | 107.8 | 264.6 | 288.3 | 299.5 | 306.2 | 313.0 | 320.1 |
| Adj. EBITDA — downside | 107.8 | 264.6 | 129.5 | 133.6 | 135.8 | 138.2 | 140.7 |
| EBITDA margin — base | 8.6% | 16.5% | 17.3% | 17.2% | 17.1% | 17.0% | 17.0% |
| EBITDA margin — downside | 8.6% | 16.5% | 9.5% | 9.4% | 9.3% | 9.2% | 9.2% |

## 2. Free cash flow

| £m | FY2026E | FY2027E | FY2028E | FY2029E | FY2030E | FY2031E |
| --- | --- | --- | --- | --- | --- | --- |
| **Base** | (230.4) | (277.6) | (266.5) | (361.6) | (388.7) | (331.6) |
| **Downside** | (230.4) | (455.4) | (451.2) | (550.9) | (582.4) | (529.9) |

Free cash flow is **negative in every year of the forecast, in both scenarios.**

## 3. Annual cash interest

| £m | FY2026E | FY2027E | FY2028E | FY2029E | FY2030E | FY2031E |
| --- | --- | --- | --- | --- | --- | --- |
| **Base** | 160.0 | 191.0 | 191.0 | 292.8 | 326.7 | 326.7 |
| **Downside** | 160.0 | 209.8 | 209.8 | 311.7 | 345.6 | 345.6 |

FY2026 is calibrated to guidance (c.£160m). The **£31.0m step-up into FY2027** is the full-year
run-rate cost of the SSTL — FY2026 carries only c.5 months of it. The second step-up in FY2029
is the refinancing of the 2029 notes at the market-implied rate.

## 4. Gross debt, net debt and leverage

| £m | FY2026E | FY2027E | FY2028E | FY2029E | FY2030E | FY2031E |
| --- | --- | --- | --- | --- | --- | --- |
| Gross debt (nominal) — base | 1,942.6 | 1,942.6 | 1,942.6 | 1,942.6 | 1,942.6 | 1,942.6 |
| **Net debt — base** | 1,645.9 | 1,899.3 | 2,101.6 | 2,278.3 | 2,422.5 | 2,501.1 |
| **Net debt / EBITDA — base** | 6.2x | 6.6x | 7.3x | 8.3x | 9.4x | 10.2x |
| Net debt / EBITDA — downside | 6.2x | 16.1x | 19.0x | 22.8x | 26.6x | 29.9x |

Leverage **rises throughout**. Starting point: 8.9x at 30 June 2026 (reported).

## 5. Interest coverage (EBITDA / cash interest)

| | FY2026E | FY2027E | FY2028E | FY2029E | FY2030E | FY2031E |
| --- | --- | --- | --- | --- | --- | --- |
| **Base** | 1.7x | 1.5x | 1.6x | 1.0x | 1.0x | 1.0x |
| **Downside** | 1.7x | 0.6x | 0.6x | 0.4x | 0.4x | 0.4x |

Base coverage never exceeds 1.7x and falls to 1.0x once the notes are refinanced.

## 6. Liquidity and runway

| £m | FY2026E | FY2027E | FY2028E | FY2029E | FY2030E | FY2031E |
| --- | --- | --- | --- | --- | --- | --- |
| **Total liquidity — base** | 307.3 | 29.7 | (236.7) | (598.4) | (987.1) | (1,318.8) |
| **Total liquidity — downside** | 407.3 | (48.0) | (499.2) | (1,050.1) | (1,632.5) | (2,162.5) |

Negative figures quantify the **funding requirement**; they are not a forecast of negative cash.

| Scenario | Crosses zero | Runway from 22-Jul-2026 |
| --- | --- | --- |
| **Base** | c. Feb-2028 | **c.19 months** |
| **Downside** | c. Nov-2027 | **c.16 months** |

Straight-line interpolation within the year — the model is annual, so treat these as indicative
quarters, not dates.

**Covenant breach: FY2027 in both scenarios**, and at every tested threshold (£75m / £100m / £125m).
The breach date does not depend on the unsourced covenant quantum.

## 7. Debt maturity profile (post-transaction, nominal £m)

| | 2026 | 2027 | 2028 | 2029 | 2030 | 2031 |
| --- | --- | --- | --- | --- | --- | --- |
| USD SSN 10.0% | – | – | – | 792.3 | – | – |
| GBP SSN 10.375% | – | – | – | 565.0 | – | – |
| SSTL (SONIA+6.75%) | – | – | – | – | – | 450.0 |
| DDTL (if drawn) | – | – | – | – | – | 100.0 |
| Inventory financing | 39.1 | – | – | – | – | – |
| Other bank loans | 6.5 | – | – | – | – | – |
| **Total** | **45.6** | – | – | **1,357.3** | – | **450.0** |

Lease liabilities (£89.7m) amortise across the period and are excluded.

The transaction **removed the 2028 RCF maturity** and pushed £450m out to 2031 — a genuine improvement
in profile. It did **not** touch the March 2029 wall, which is **c.4x** the c.£340m pro forma liquidity.

## 8. Refinancing feasibility for the 2029 notes

FY2030, first full year post-refinancing, base case:

| Refinancing rate | FY30 interest | Coverage | FY30 FCF | Peak funding gap |
| --- | --- | --- | --- | --- |
| **0.00% (hypothetical)** | 53.1 | 5.89x | **(115.1)** | 566.3 |
| 7.50% | 154.9 | 2.02x | (216.9) | 846.2 |
| **10.375% (existing coupon)** | 193.9 | **1.61x** | (255.9) | 953.5 |
| 12.50% | 222.8 | 1.41x | (284.8) | 1,032.8 |
| 15.00% | 256.7 | 1.22x | (318.7) | 1,126.2 |
| **20.16% (market-implied)** | 326.7 | **0.96x** | (388.7) | 1,318.8 |

**The decisive result: at a 0% refinancing cost — an impossible best case — FY2030 free cash flow is
still £(115.1)m.** Refinancing terms are not the binding constraint. No achievable coupon makes this
capital structure work.

Refinancing the 2029 notes is therefore a question about **the business**, not about credit markets.
Even refinancing at the *existing* 10.375% coupon leaves 1.6x coverage and deeply negative cash flow —
and the notes were marked at **79.6% of par** at 30 June 2026, implying a 20.16% yield, so the existing
coupon is not available.

## 9. Is further funding required before 2031?

**Yes, in both scenarios.**

| | First required | Peak funding gap |
| --- | --- | --- |
| Base | FY2028 | **£1,318.8m** |
| Downside | FY2027 | **£2,162.5m** |

The £100m undrawn DDTL is the only committed capacity remaining and defers the base-case breach by
roughly one quarter. Beyond it there is no undrawn facility — the RCF and YTC facility were both
cancelled at closing.

## 10. THE HEADLINE OUTPUT — self-sustaining, or short of liquidity again?

The brief's single most important output. **Base case, before any interest:**

| £m | FY2026E | FY2027E | FY2028E | FY2029E | FY2030E | FY2031E |
| --- | --- | --- | --- | --- | --- | --- |
| Adjusted EBITDA | 264.6 | 288.3 | 299.5 | 306.2 | 313.0 | 320.1 |
| Capex | (300.0) | (350.0) | (350.0) | (350.0) | (350.0) | (300.0) |
| **EBITDA less capex** | **(35.4)** | **(61.7)** | **(50.5)** | **(43.8)** | **(37.0)** | **20.1** |

**EBITDA does not cover capex alone until FY2031 — and then by only £20m, before a penny of interest.**

### What would have to be true for FY2031 free cash flow to reach zero?

| Requirement | Needed | vs modelled |
| --- | --- | --- |
| Adjusted EBITDA | £651.7m | vs £320.1m — **more than double** |
| Implied revenue | c.£2,823m | vs £1,888m — **c.+50%** |
| Implied wholesale volumes | c.9,870 units | vs 6,600 modelled; **5,448 actual in FY2025** |
| Even refinancing at 10.375% | still needs c.£519m EBITDA | c.+62% |

For context, AML's **peak historic adjusted EBITDA was £271m (FY2024)** and peak recent volumes were
6,030 (FY2024). The self-sustaining threshold requires roughly **1.6x its best-ever volume year**.

### Answer

**Aston Martin does not become self-sustaining within the forecast horizon. It runs short of liquidity
first — c.February 2028 in the base case, c.November 2027 in the downside.**

---

## 11. Alternative financing options (qualitative, per the brief)

Not modelled in detail. Anchored to model output where honest.

| Alternative | Advantage | Disadvantage | What our numbers say |
| --- | --- | --- | --- |
| **Rights issue** | Reduces leverage and interest; only option that fixes solvency, not just liquidity | Severe dilution; Yew Tree Consortium already stretched | Would need to be c.£1.3bn to close the base-case gap — roughly equal to FY2025 revenue (£1,258m). Very hard to place |
| **Convertible bond** | Lower cash interest preserves near-term liquidity | Future dilution; complex to price on a distressed credit | Saves cash interest but does not touch the £62m FY2027 EBITDA-less-capex deficit. Defers, does not fix |
| **Public secured bond** | Fixed rate, longer tenor, no floating exposure | Likely unplaceable given the credit | The existing notes were marked at **79.6% of par (20.16% yield)** at 30 June 2026 — the public market was effectively shut |
| **Private credit (actual deal)** | Fast, certain execution; flexible terms; met an urgent need | High floating cost; more secured claims; subordinates existing bondholders; encumbers the balance sheet | Delivered c.£340m pro forma liquidity and removed the 2028 RCF maturity, at **+£31m** annual interest. Bought **c.19 months** |
| **Mixed debt / equity** | Balances liquidity with genuine deleveraging | More complex; still dilutive; slower to execute | The only structure that addresses **both** the liquidity gap and the structural EBITDA-capex deficit |

**Assessment.** Given the timeline pressure, the state of the public market for AML paper, and a
shareholder base already stretched, the private-credit package was plausibly **the only executable
option** in the window available. That makes it defensible as an execution decision. But it solved a
**liquidity** problem, and the model says the problem is **structural** — so it bought time at high
cost rather than fixing the capital structure.

---

## 12. Where the evidence points

Against the brief's three permitted outcomes, the model points to the **third**:

> **Financing merely delayed restructuring** — still too leveraged; equity raise, asset sale, debt
> restructuring or further shareholder support likely required.

Supporting evidence, in order of strength:

1. At a **0% refinancing cost**, FY2030 free cash flow is still £(115)m — no credit-market outcome fixes it.
2. **EBITDA does not cover capex alone until FY2031**, before any interest.
3. Liquidity runway is **c.19 months**; covenant breach in **FY2027** at every tested threshold.
4. Leverage **rises** to 10.2x; coverage falls to 1.0x.
5. Reaching self-sustaining requires **c.+50% revenue** versus a base already capped at AML's best-ever margin.

⚠️ The honest counterweight, to be stated rather than buried: the conclusion rests on **FY2027–31
drivers that AML has not guided**, and on a capex programme whose £1.7bn total is sourced but whose
phasing is ours. Re-phasing capex moves the dates by up to a quarter but **not** the conclusion — the
programme total is guidance. See `SOURCES.md` §8d.

*The final recommendation is Phase 5. This section records where the evidence points, not a decision.*
