# Aston Martin Lagonda — £550m Private Credit Financing: Solved, or Postponed?

A credit and capital-structure analysis of Aston Martin Lagonda's July 2026 £550m private-credit
financing. **Not a valuation exercise** — no comparables, no precedent transactions, no football field.

> **Central question:** did the financing solve Aston Martin's liquidity problem, or merely postpone
> the need for further deleveraging or equity funding?

## Conclusion

**The financing delayed a restructuring rather than avoiding one.** It solved a liquidity problem; the
analysis indicates the problem is structural. A further capital event — additional debt, an equity
injection, an asset sale, or further shareholder support — is expected before the March 2029 maturity.

The three findings that carry the analysis, chosen because they depend on the fewest assumptions:

| Finding | Why it matters |
| --- | --- |
| **At a 0% cost of refinancing, FY2030 free cash flow is still £(115)m** | No achievable coupon fixes the capital structure. Refinancing is a question about the business, not credit markets |
| **EBITDA does not cover capex alone until FY2031** — and then by £20m, before any interest | The deficit is structural, and this holds without any interest-rate assumption |
| **The 2029 notes were marked at 79.6% of par at 30 June 2026** (down from 92.6% in December, before the deal) | The market had already priced a restructuring |

Liquidity runway from the financing: **c.19 months** (base), **c.16 months** (downside). Covenant breach
in **FY2027** in both cases and at every tested threshold.

## Deliverables

| File | What it is |
| --- | --- |
| `AML_Financing_Deck.pptx` / `.pdf` | 6-slide capital-markets summary deck, with speaker notes |
| `AML_Recommendation_Memo.docx` / `.pdf` | One-page investment-banking style recommendation memo |
| `AML_Liquidity_Model.xlsx` | The model — 10 tabs, live formulas, base and downside scenarios |
| `OUTPUTS.md` | Every output the brief requires, both scenarios, plus the alternatives table |
| `SOURCES.md` | Sourcing record — every figure with its primary-source quote, plus all derivations |
| `PROGRESS.md` | Working log: decisions, verification, open items, defend-list |

## The model

`AML_Liquidity_Model.xlsx` — the **Assumptions** tab is the only one to edit. `Assumptions!B3` switches
between base (1) and downside (2). Rebuild from scratch with `build_model.ps1` (Excel COM, idempotent).

| Tab | Contents |
| --- | --- |
| README | Purpose, colour code, key flags |
| Assumptions | All inputs. Base block, downside block, active drivers, switches |
| Debt | Pre/post capital structure, interest schedule, reconciliation to reported net debt |
| Model | P&L, free cash flow, net debt and liquidity roll-forward FY2025A–FY2031E |
| Outputs | Credit metrics dashboard, covenant sensitivity, scenario comparison |
| YTM | Implied yield derivation on the 2029 notes from the disclosed market price |
| Attribution | What drives the leverage and coverage reversal |
| Refinancing | 2029 refinancing feasibility, self-sustaining test, liquidity runway |
| Reconciliation | Why the model exhausts liquidity earlier than management's stress test implies |
| Maturity | Debt maturity profile |

### Verification built in

- Nominal debt less unamortised fees ties **exactly** to reported gross debt (£1,661.1m)
- Pre-transaction liquidity ties **exactly** to the reported £145.2m
- Post-transaction liquidity reproduces the disclosed c.£340m
- FY2025A adjusted EBIT rebuilds from drivers to the reported £(189.2)m
- All three reported leverage ratios recompute (8.9x / 12.8x / 6.7x)

## Sourcing discipline

**Every figure in the model traces to an Aston Martin primary document** — H1 2026 interim results,
Q1 2026 results, or the FY2025 Annual Report — read from locally extracted text, never from a web
summariser. `SOURCES.md` carries a **verbatim quote and document location for every figure**, so the
audit trail is complete without the source files themselves.

The source PDFs are **not redistributed in this repository** — they are Aston Martin's own copyrighted
documents. All are freely available from
[Aston Martin investor relations](https://www.astonmartin.com/en/corporate/investors). To rebuild the
local working copy, download them into `source_docs/`:

| Document | Where |
| --- | --- |
| H1 2026 interim results (+ presentation, transcript) | Results and presentations |
| Q1 2026 results | Results and presentations |
| Annual Report 2025 | Annual report |

Two external inputs, both labelled: the SONIA base rate (Bank of England) and the implied yield on the
2029 notes (computed from Aston Martin's own disclosed fair values).

Press reporting on the transaction's structure appears **once** in the deck and **once** in the memo as
labelled market context. It is **not an input to any figure and no conclusion depends on it.**

## Caveats that travel with the numbers

1. **The 20.16% refinancing rate is a market-implied distress yield**, not a forecast new-issue coupon.
   The conclusion does not depend on it — the 0% test makes the same point.
2. **The downside is a no-response case.** It excludes the mitigating actions management states it
   would take, matching the construction of management's own disclosed reverse stress test.
3. **FY2027–31 drivers are unsourced judgement**, capped at Aston Martin's own FY2024 best margin
   (17.1%). The conclusion is sensitive to the base case, not just the stress — quantified in
   `SOURCES.md` §8d.

## Reproducing

Requires Windows with Excel, PowerPoint and Word (COM automation).

```powershell
.\build_model.ps1    # rebuilds the workbook
.\build_deck.ps1     # rebuilds the deck
.\build_memo.ps1     # rebuilds the memo
```

---

*Prepared as a personal analytical project from public disclosures. Not investment advice.*
*`DRAFT_deck_and_memo.md` is the reviewed content draft, retained as an audit trail.*
