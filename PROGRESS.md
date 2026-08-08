# PROGRESS — Aston Martin £550m Private Credit Financing Analysis

**To resume after a crash/restart: read this file top to bottom, then do "Next action" at the bottom.**

Last updated: 2026-08-07 — **Phase 2 (model build) COMPLETE. Awaiting sign-off. Phase 3 NOT authorised.**

> ⛔ **STOP GATE:** Do not begin Phase 3 (downside scenario) until the user explicitly says go.
> Downside parameters must be proposed and signed off, never chosen silently.

---

## Project

CV / interview-prep project (IB, corporate finance, TS, PE recruiting).
**Credit and capital-structure analysis — NOT a valuation exercise.** No comps, no precedent
transactions, no football field.

**Central question:** Did Aston Martin's July 2026 £550m private-credit financing solve its liquidity
problem, or merely postpone further deleveraging or an equity raise?

- Brief: `aston_martin_project_plan.md.pdf` (in this folder; it is a PDF despite the `.md` in the name)
- Approved plan: `C:\Users\harry\.claude\plans\read-the-aston-martin-majestic-sphinx.md`
- **Verified figures + verbatim quotes: `SOURCES.md`** ← single source of truth for all model inputs

## Phases (sign-off required after each)

| # | Phase | Status |
| --- | --- | --- |
| 1 | Data sourcing & verification | ✅ **COMPLETE — SIGNED OFF 2026-08-07** |
| 2 | Model build (Excel) | ✅ **COMPLETE — SIGNED OFF 2026-08-07** |
| 3 | Scenarios (base + one downside) | ✅ **COMPLETE — SIGNED OFF 2026-08-07** |
| 4 | Outputs | ✅ **COMPLETE — awaiting sign-off** |
| 5 | Deck (4–6 slides) + one-page memo | ✅ **COMPLETE — final files built** |
| 6 | GitHub repo | 🟡 **Local repo initialised and committed. Not pushed — remote TBD** |

---

## Decisions made

| Date | Decision | Detail |
| --- | --- | --- |
| 2026-08-07 | Project folder | `C:\Users\harry\aston-martin-financing-analysis` |
| 2026-08-07 | Sourcing rule | astonmartin.com for **all model inputs**. Tier-1 press only for deal context AML does not disclose, flagged, never a model input. |
| 2026-08-07 | Excel approach | Real `.xlsx` via Excel COM — live formulas, linked tabs, scenario switch |
| 2026-08-07 | Filing type correction | AML is LSE-listed: UK interim/annual reports + RNS. The brief's "10-K/10-Q" does not exist. |
| 2026-08-07 | Figure discipline | All numbers read by Claude from locally extracted PDF text, never from a web summariser |
| 2026-08-07 | `data/core_figures.md` deferred | SOURCES.md already is the verified figure set; a second copy would risk divergence. The model input register gets built in Phase 2 directly from SOURCES.md. **Flagged to user.** |

## Environment (verified 2026-08-07)

- No Python, no poppler/pdftotext → the `Read` tool **cannot** open PDFs.
- **PDF text extraction = Word COM** (see the extraction command pattern in git history of this file / re-derive: open read-only, `$doc.Content.Text`, write UTF8 `.txt`).
- Extracted `.txt` files are one continuous line → grep by line fails. Search with
  `$t.IndexOf(term)` + `Substring` context windows instead.
- Excel COM available (v16.0).
- `git` NOT installed → blocks phase 6 only.
- Reuters.com and FT.com block our user agent. Bloomberg.com is reachable via WebSearch.
- IR site: `https://www.astonmartin.com/en/corporate/investors` (the site's own 301 to `www2.` is dead).

---

## Step 1 result: all four consistency checks PASS

| Check | Result |
| --- | --- |
| Gross debt − cash = net debt | 1,661.1 − 114.9 − 1.5 = **1,544.7** ✅ |
| Leverage recomputes | 1,544.7 ÷ 173.8 = **8.89x** → reported 8.9x ✅ (also 12.8x and 6.7x comparatives ✅) |
| H1 FCF = Q1 + Q2 | 116.8 + 80.8 = **197.6** ✅ |
| Loan note tranches = bridge | 786.3 + 461.6 + 97.3 = **1,345.2** ✅ |

Fifth check — pro forma liquidity — **does not fully reconcile** (see open questions).

## Every brief figure verified

| Brief said | Primary source says | Verdict |
| --- | --- | --- |
| £1.5bn net debt | £1,544.7m | ✅ |
| £198m H1 FCF outflow | £197.6m | ✅ |
| £109m adjusted EBIT loss | £(108.9)m | ✅ |
| £145m total liquidity | £145.2m | ✅ |
| 8.9x adjusted net leverage | 8.9x | ✅ |
| ~$1.05bn secured notes @ 10% | $1,050.0m @ 10.0%, Mar 2029 | ✅ |
| ~£565m secured notes @ 10.375% | £565m @ 10.375%, Mar 2029 — **two tranches (£465m + £100m)** | ✅ w/ nuance |
| £163m drawn on £170m RCF (fully drawn) | £163.0m cash + £5.9m LC reserve = fully utilised | ✅ both true |
| £20m Yew Tree facility | **£50m** facility, **£20m drawn** | ⚠️ brief understated facility size |
| £450m SSTL + £100m DDTL | ✅ | ✅ |
| SONIA + 6.75%, July 2031 | ✅ | ✅ |
| £100m junior debt capacity | ✅ | ✅ |
| HPS lead lender | ✅ **confirmed in primary doc, not just press** | ✅ |
| Net cash interest ~£160m (from ~£150m) | c.£160m (previously c.£150m) | ✅ |
| Pro forma liquidity ~£340m | c.£340m | ✅ |

## New findings not in the brief (material to the analysis)

1. **2029 notes marked at 79.6% of par** at 30 Jun 2026 (FV £1,080.0m vs £1,357.3m nominal), down from
   92.6% at 31 Dec 2025 — *before* the July financing. Direct market read on refinancing risk.
2. **Inventory repurchase £39.1m repayable October 2026** (£40.0m gross) — near-term cash call.
3. **SSTL covenant = minimum liquidity, tested monthly from Aug 2026** (quantum undisclosed). The RCF
   leverage covenant is no longer tested. Management reverse stress test: core volumes −55% exhausts
   liquidity, −25% breaches covenants.
4. Implied **FX 1.3253 USD/GBP** at 30 Jun 2026 (derived: $1,050.0m ÷ £792.3m) — no external FX source needed.
5. FY2025 base year: revenue £1,257.7m, adj EBITDA £108.1m, adj EBIT £(189.2)m, volumes 5,448.
6. Press (unverified): deal reportedly structured as a **drop-down**; creditor group (Arini, BlackRock,
   Sculptor) signed a co-op agreement and sent a legal letter. BlackRock owns HPS *and* appears among
   the objectors.

## Open questions — ALL THREE RESOLVED 2026-08-07

**1. Pro forma liquidity bridge — RESOLVED (user sign-off, flagged assumption).**
Modelled as a sensitivity band, not a point estimate: **£30m / £40m (central) / £50m** of transaction
costs → pro forma liquidity c.£352m / c.£342m / c.£332m. Status: unverified, pending FY2026 Annual
Report financing cash flows. Must stay visibly flagged in the model notes. See `SOURCES.md` §8a/A1.

**1b. £100m DDTL — RESOLVED and UPGRADED to primary-sourced.**
The CFO confirmed on the H1 call: *"the delayed draw term loan element of that is not included in that
liquidity number."* Treatment: **memo item only**, committed-but-undrawn capacity on top of the c.£340m.
Never added to headline liquidity; carries its own drawn/undrawn sensitivity and footnote. See §8a/A2.

**2. SSTL covenants — RESOLVED from primary source; the RNS was not needed.**
The interim discloses a **minimum liquidity covenant, tested monthly from August 2026**, and that the
RCF leverage covenant is no longer tested. ⚠️ The **quantum is not disclosed** — recorded as a
limitation (per instruction, searching stopped rather than continuing). The model will show liquidity
against a user-selectable threshold, clearly labelled as unsourced.
**Bonus find:** management's own reverse stress test — core volumes must fall **>55%** from forecast to
exhaust liquidity and **>25%** to breach covenants. This is the natural calibration anchor for the
Phase 3 downside case (to be proposed for sign-off, not adopted silently).

**3. Drop-down / bondholder friction — RESOLVED (bounded role).**
Permitted: **one** labelled "market context" callout each in memo and deck, plus colour for the "why
private credit over alternatives" discussion. Prohibited: any model input, any number, any conclusion.
Label: *"Reported bondholder friction over asset ranking, per Bloomberg reporting — not confirmed by
AML."* The analytical weight is carried instead by the primary-sourced fact that the 2029 notes were
marked at **79.6% of par**. Full rule recorded in `SOURCES.md` §9.

## Deferred decisions (raise later, do not assume)

- Downside-case parameters (deliveries, margin, SONIA path) — propose in Phase 3, get sign-off.
- SONIA forward curve source — not on astonmartin.com; needs its own sourcing decision.
- `git` install — needed before Phase 6.

---

## Files created

| File | Purpose |
| --- | --- |
| `PROGRESS.md` | This file |
| `SOURCES.md` | **All verified figures + verbatim primary-source quotes** |
| `source_docs/*.pdf` | 5 primary documents downloaded from astonmartin.com |
| `source_docs/*.txt` | Word-COM text extractions of each |

---

---

## PHASE 2 RESULT — model built and internally verified

**File: `AML_Liquidity_Model.xlsx`** (rebuild any time with `build_model.ps1` — Excel COM, idempotent).
Eight tabs: README · Assumptions · Debt · Model · Outputs · **YTM** · **Attribution** · Maturity. Live
formulas throughout; the Assumptions tab is the only one to edit.

> ⚠️ **PowerShell/COM gotcha:** writing an Int32 to `Range().Value2` throws "Specified cast is not
> valid" and, in a loop, **fails silently for those cells while others succeed** — which produced one
> invalid attribution run before it was caught. Always cast: `.Value2 = [double]$x`. Sanity-check any
> scenario walk by confirming the final step reconciles to the live workbook.

> ⚠️ **Excel COM gotcha:** force-killing EXCEL.EXE mid-run wedges the COM server and every subsequent
> `Workbooks.Open` fails with RPC_E_CALL_REJECTED. The workbook itself is fine. Either wait and retry
> in a fresh process, or verify values by reading the cached `<v>` elements straight out of the
> xlsx XML (it is a zip) — that path is COM-free and was used to validate this build.

### Validation checks — all pass

| Check | Result |
| --- | --- |
| Nominal debt less unamortised fees = reported gross debt | 1,675.6 − 14.5 = **1,661.1** ✅ exact |
| Pre-transaction liquidity | model **145.2** vs disclosed 145.2 ✅ exact |
| Post-transaction pro forma liquidity | model **341.9** vs disclosed c.340 ✅ (£40m central cost case) |
| FY2025A adjusted EBIT rebuilt from drivers | **(189.2)** vs reported (189.2) ✅ |
| FY2025A revenue / EBITDA rebuilt from drivers | 1,257.9 / 107.8 vs 1,257.7 / 108.1 ✅ rounding only |
| Switches (DDTL on/off, refi on/off, cost case) | all flow through correctly ✅ |

### Base-case output — REVISED 2026-08-07 after Phase 2 review

| GBP m | FY26E | FY27E | FY28E | FY29E | FY30E | FY31E |
| --- | --- | --- | --- | --- | --- | --- |
| Revenue | 1,598.9 | 1,671.4 | 1,745.1 | 1,792.0 | 1,839.5 | 1,887.6 |
| Adjusted EBITDA | 264.6 | 288.3 | 299.5 | 306.2 | 313.0 | 320.1 |
| Adj EBITDA margin | 16.5% | 17.3% | 17.2% | 17.1% | 17.0% | 17.0% |
| Net cash interest | (160.0) | (191.0) | (191.0) | (292.8) | (326.7) | (326.7) |
| Free cash flow | (230.4) | (277.6) | (266.5) | (361.6) | (388.7) | (331.6) |
| Total liquidity | 307.3 | 29.7 | (236.7) | (598.4) | (987.1) | (1,318.8) |
| Funding gap | – | – | 236.7 | 598.4 | 987.1 | 1,318.8 |
| Net debt / EBITDA | 6.2x | 6.6x | 7.3x | 8.3x | 9.4x | 10.2x |
| EBITDA / cash interest | 1.7x | 1.5x | 1.6x | 1.0x | 1.0x | 1.0x |

**Headline:** covenant breach **FY2027**, liquidity exhausted **FY2028**, peak funding gap
**GBP 1,318.8m** by FY2031, 2029 refinancing requirement **GBP 1,357.3m**. FY2027 cash interest is
**GBP 31.0m above** FY2026 — the full-year run-rate cost of the deal.

⚠️ **Earlier attribution corrected (2026-08-07).** An earlier note here said the FY2029 step-change was
driven by the refinancing rate. That was wrong on the leverage measure. Formal attribution (below, and
on the **Attribution** tab) shows the *leverage* reversal is **78% margin cap / 23% refi rate**; only
the *coverage* collapse is an even split.

⚠️ **REFI RATE CAVEAT — must travel with the number into the memo and deck.** 20.16% is the
market-implied **distress yield** on the 2029 notes at 30 June 2026, not a forecast new-issue coupon;
it embeds default probability. Placed at Assumptions I29 · Debt H38 · Outputs A4 · README B29 · YTM A19.
Full rule in `SOURCES.md` §8b/B2.

### Attribution of the reversal (Attribution tab; S4 reconciles to the live model)

| Metric | Total swing | Margin cap | SONIA | FY26 calib. | Refi rate |
| --- | --- | --- | --- | --- | --- |
| FY2031 leverage | +4.62x | **+3.61x (78%)** | −0.02x | −0.03x | +1.06x (23%) |
| FY2029 coverage | −0.94x | −0.47x (50%) | +0.01x | 0.00x | −0.48x (51%) |
| FY2031 coverage | −1.23x | −0.64x (52%) | +0.01x | 0.00x | −0.60x (49%) |
| Peak funding gap | +£758.8m | **+£435.3m (57%)** | −£6.9m | −£11.5m | +£341.9m (45%) |

Not a compounding of unexamined assumptions: two deliberate corrections did the work, each replacing a
weaker input with a better-grounded one. **Robustness — at margin cap alone with refi still at 11%,
leverage already reaches 9.17x and the funding gap £995m. The conclusion does not depend on the
distress yield.**

### All five Phase 2 review points — CLOSED

| # | Point | Resolution |
| --- | --- | --- |
| 1 | Cap EBITDA margin at c.17.1% | ✅ Done. Searched all AML docs for a quantified target — **none exists**, so 17.1% (FY2024 high) is the ceiling. Path peaks 17.3%, settles 17.0%. Basis stated in `SOURCES.md` §8b/B4 and on the Assumptions tab. |
| 2 | Calibrate FY2026 interest to c.£160m | ✅ Done. Explicit **−£11.6m** line on Debt tab row 42; FY2026 now ties to guidance exactly (delta 0.00). Accrual retained FY2027+. Accrual-vs-cash note on the tab. |
| 3 | Source SONIA + check forward curve | ✅ Done. **3.72%** = BoE Bank Rate 3.75% (Jul-2026) less c.3bp. ⚠️ Forward curve **not retrievable** (BoE returns 403) — checked, not assumed. Materiality low: ±100bp = ±£4.5m. |
| 4 | Back out YTM instead of guessing 11% | ✅ Done. Blended implied YTM **20.16%** (USD 20.02 / GBP465 20.14 / GBP100 21.44). Full workings on the new **YTM** tab. Distress-yield caveat recorded. |
| 5 | Covenant threshold sensitivity | ✅ Done. New block on Outputs tab. **Breach year is FY2027 at £75m, £100m and £125m** — the conclusion does not hinge on the unsourced input. |

### Validation after rebuild

Debt reconciliation still **TIES** (£1,661.1m). Post-transaction liquidity **£341.9m** vs disclosed
c.£340m. FY2026 net cash interest **£160.0m** vs guidance £160m, delta 0.00. FY2025A adjusted EBITDA
margin rebuilds to **8.6%** — matches the reported figure.

---

### Phase 2 closing note (superseded by Phase 3 below)

All five Phase 2 review points closed and signed off. One limitation carried forward, not a blocker:
the BoE SONIA **forward curve** could not be retrieved (403); SONIA is held flat at the sourced spot
rate of 3.72% in the base case.

### PHASE 3 — BUILT 2026-08-07. Awaiting sign-off on the result.

All four decisions signed off: **A** −23% volumes (traceable translation), **B** exclude mitigating
actions, **C** SONIA 5.25% as a labelled judgement, **D** DDTL drawn.

The workbook now holds **both cases in one file**, driven by `Assumptions!B3` (1 = Base, 2 = Downside).
Active driver rows 13–22 and SONIA row 26 are formulas selecting between the **BASE INPUTS** block
(rows 42–52) and the **DOWNSIDE INPUTS** block (rows 57–67). DDTL (B34) switches automatically.
Derivation detail in `SOURCES.md` §8c.

| Metric | Base | Downside |
| --- | --- | --- |
| FY2027 revenue | £1,671.4m | £1,357.8m (−18.8%) |
| FY2027 adjusted EBITDA | £288.3m | £129.5m (−55%) |
| FY2027 EBITDA margin | 17.3% | 9.5% |
| FY2027 free cash flow | £(277.6)m | £(455.4)m |
| Covenant breach | FY2027 | FY2027 |
| **Liquidity exhausted** | **FY2028** | **FY2027** |
| Peak funding gap | £1,318.8m | £2,162.5m |
| FY2031 net debt / EBITDA | 10.2x | 29.9x |
| FY2029 EBITDA / cash interest | 1.0x | 0.4x |

**Verification:** switch integrity confirmed both ways (B3=1 → FY27 revenue £1,671.4m, SONIA 3.72%,
DDTL 0; B3=2 → £1,357.8m, 5.25%, DDTL 1). Debt reconciliation still TIES; base post-transaction
liquidity £341.9m vs disclosed c.£340m. Workbook left saved on **BASE**.

⚠️ The Debt tab's "CHECK vs disclosed" is calibrated to the base case (DDTL undrawn), as AML disclosed
it. It will not tie in the downside by construction — noted on the tab itself.

### Superseded proposal (kept for the audit trail)

Single combined case per the brief (lower deliveries + weaker margins + higher SONIA together, not in
isolation). Anchored on management's disclosed reverse stress test (`SOURCES.md` §8): core volumes
(DBX and GT/Sports) −25% from forecast breaches covenants, −55% exhausts liquidity.

| # | Parameter | Base | **Proposed downside** | Basis |
| --- | --- | --- | --- | --- |
| 1 | Wholesale volumes (FY2027+) | 6,100→6,600 | **−23% vs base** | = management's −25% **core** stress with Specials held flat. Specials ≈8.5% of units (c.500 Valhalla of c.5,900 FY26), so −25% core ≈ −23% total |
| 2 | Blended ASP | 274→286 | **+5.5% vs base** | Mix effect only — Specials become a larger share of a smaller book. Net revenue ≈ **−18.5%** |
| 3 | Gross margin | 35.5% | **32.0%** (−350bps) | Between the FY2025 trough (29.4%, sourced) and the H1-26 run-rate (33.8%, sourced) |
| 4 | Adjusted opex | 305→350 | **unchanged** | Fixed cost base; this is the operating-deleverage effect |
| 5 | Capex | 350 | **unchanged** | Management's reverse stress test explicitly excludes mitigating actions |
| 6 | SONIA | 3.72% | **5.25%** (+153bps) | ⚠️ JUDGEMENT — no forward curve obtainable (BoE 403) |
| 7 | DDTL | undrawn | **drawn (£100m)** | Under stress AML would draw it; adds debt at SONIA+6.75% |
| 8 | Refi rate | 20.16% | **unchanged** | Already a distress yield; stressing further would double-count |
| 9 | Transaction costs / covenant | £40m / £100m | **unchanged** | Isolates the operating stress |

All four decisions above were signed off as proposed and are now built.

**Note on base-case consistency:** the base case breaches covenant in FY2027 with year-end liquidity of
£29.7m, while management expects covenant compliance through 30 Sep 2027. These are broadly reconcilable
— the going-concern period ends before FY2027 year-end and management's £340m starting liquidity is the
same. Not a contradiction, but worth stating in the memo.

---

## DEFEND-LIST ITEM #1 — resolved 2026-08-07

**Challenge:** the downside uses −23% volumes (milder than management's −25% covenant marker) yet fully
exhausts liquidity, when management's disclosure implies −55% is needed to exhaust. Why?

**Answer: it is not the margin lever — it is the base case.**

- Volume-only −25% (margin and SONIA at base) **already exhausts** liquidity in FY2027 (−£81.8m). The
  margin cut adds a further −£47.5m but is not required. SONIA contributes −£6.9m. DDTL adds back +£88.0m.
- Volume-only ladder: our model tips into FY2027 exhaustion at just **−6.6% core volumes** vs
  management's −55%. Base FY2027 year-end liquidity is only **£29.7m** — on the cliff edge.
- Structural cause: FY2027 interest £191m + capex £350m = **£541m committed against £288m EBITDA**.
- **Not** a capex-phasing artefact: re-phasing within the sourced £1.7bn total moves FY2027 by up to
  £100m but never changes the exhaustion year.
- Four non-like-for-like differences: different starting forecast (the big one), different horizon
  (Sep-27 vs Dec-27), volume-only vs multi-lever, mitigating actions.

**Defensible line:** *"On a base case capped at AML's own historic best margin and carrying its own
guided capex programme, the capital structure is already cash-negative before any stress. That is the
finding. The downside simply confirms there is no room for error."*

⚠️ **Caveat to state, not hide:** the conclusion is sensitive to the **base case**, not just the
downside. Management's forecast is undisclosed, so the divergence can be bounded, not reconciled.

Full workings: **Reconciliation** tab (Tests 1–3) and `SOURCES.md` §8d.

---

## PHASE 4 — COMPLETE 2026-08-07

**Deliverable: `OUTPUTS.md`** — every output the brief requires, both scenarios, plus the qualitative
alternatives table. Workbook gained a **Refinancing** tab (2029 feasibility, self-sustaining test,
runway) and a liquidity-runway line on Outputs.

### The three findings that carry the analysis

1. **At a 0% refinancing cost, FY2030 free cash flow is still £(115.1)m.** No achievable coupon fixes
   the capital structure — refinancing is a question about the business, not credit markets. Even at
   the existing 10.375% coupon, coverage is 1.6x and FCF £(255.9)m.
2. **EBITDA does not cover capex alone until FY2031** (£20.1m), before any interest. FY2027 is £(61.7)m.
3. **Liquidity runway c.19 months** (base, crosses zero c.Feb-2028); c.16 months downside (c.Nov-2027).

Self-sustaining threshold: FY2031 EBITDA of **£651.7m** (vs £320.1m modelled), implying c.£2,823m
revenue and c.9,870 units — against a **FY2024 peak of £271m EBITDA on 6,030 units**.

### Where the evidence points

The brief's **third** outcome: *financing merely delayed restructuring*. Recorded as evidence, not a
decision — the recommendation is Phase 5.

### Full output set

Revenue/EBITDA · FCF · cash interest · gross and net debt · leverage · coverage · liquidity and runway ·
maturity profile · 2029 refinancing feasibility · further-funding requirement · self-sustaining point ·
alternatives table. All in `OUTPUTS.md`.

---

## PHASE 5 — DRAFT ISSUED FOR REVIEW 2026-08-08

**`DRAFT_deck_and_memo.md`** — content and framing only, formatting deliberately deferred.

6-slide deck: (1) Executive conclusion · (2) Why needed + transaction terms · (3) Pre/post capital
structure + maturity profile · (4) **The structural test — analytical core** · (5) Scenario output ·
(6) Alternatives + recommendation. Plus a c.700-word one-page memo.

**Robustness hierarchy applied:** slide 4 leads on the assumption-light findings (0% refi test, EBITDA
less capex, 79.6% of par). Leverage and coverage ratios appear on slide 5 as support only. Nothing
load-bearing rests on the refi rate, SONIA path or covenant threshold.

**Recommendation made:** outcome 3 — *delayed a restructuring rather than avoiding one* — with the
reasoning shown, and with execution quality separated from problem-solving (the deal is defensible as
execution; it does not work as a solution). "What would have to be true" for the softer verdict is
stated explicitly: FY2031 EBITDA of £651.7m vs £320m modelled.

**All three caveats travel on-slide** (distress yield on slide 4, no-response on slide 5, press context
once each in deck and memo).

**Four open questions posed for review:** slide count (6 vs merging to 5), recommendation strength,
memo length, market-context callout placement.

---

## PHASE 5 — COMPLETE 2026-08-08

Four review decisions applied: **six slides kept** (maturity exhibit retained), **recommendation
specifies the forms** a capital event could take, **memo held to one page**, **market-context callout
moved to slide 6** (alternatives).

| Deliverable | Status |
| --- | --- |
| `AML_Financing_Deck.pptx` / `.pdf` | 6 slides, speaker notes on every slide. Built by `build_deck.ps1` |
| `AML_Recommendation_Memo.docx` / `.pdf` | **1 page, 654 words** (verified via Word statistics). Built by `build_memo.ps1` |
| `README.md` | Repo index and conclusion summary |

**Visual QA:** all six slides exported to PNG and inspected. Two defects found and fixed — bullets were
rendering as a numbered list, and slide 6's Assessment column was right-aligned. Both corrected and
re-verified.

**Deck structure:** (1) Executive conclusion · (2) Why needed + terms · (3) What changed / maturity
profile · (4) **The structural test** · (5) Scenario output · (6) Alternatives + recommendation.
Robustness hierarchy applied — slide 4 carries the argument, slide 5 supports it.

**All three caveats travel on-slide**: distress yield (slide 4), no-response downside (slide 5), press
context (slide 6 only, once, plus once in the memo).

---

## PHASE 6 — LOCAL REPO DONE 2026-08-08

Git **2.55.0.windows.3** installed via winget (`Git.Git`). Repo initialised on branch `main`.

- **Commit `c47da73`** — 25 files, 2,741 insertions. Working tree clean. **No remote, nothing pushed.**
- Identity set **repo-local only** (`Harry <harrydowney123@gmail.com>`) — global git config left
  untouched. Change with `git config user.name "..."` before pushing if a full name is wanted.
- `.gitignore` covers Office lock files (`~$*`) and OS cruft only.

### Pre-push decisions — ALL FOUR APPLIED 2026-08-08

| Item | Decision | Status |
| --- | --- | --- |
| `source_docs/` | Exclude — AML copyrighted PDFs, 95% of repo size | ✅ untracked via `.gitignore`; **still on disk locally** |
| `aston_martin_project_plan.md.pdf` | Exclude — personal brief | ✅ untracked; **still on disk locally** |
| `DRAFT_deck_and_memo.md` | Keep — shows the review process | ✅ tracked |
| Commit author | **Harry Downey** `<harrydowney1234@gmail.com>` | ✅ set repo-local; global config untouched |

**Repo went from 16.5MB / 25 files to 0.65MB / 14 files.**

`README.md` carries the IR download list for rebuilding `source_docs/` locally; `SOURCES.md` notes the
PDFs are not redistributed and that every figure still has a verbatim quote.

### Remote — configured, push pending user action

**Remote added and verified:**
`origin  https://github.com/harrydowney1234-dot/aston-martin-financing-analysis.git`

⛔ **The push cannot be run from inside this session.** There is no TTY (`/dev/tty: No such device`)
and `GIT_TERMINAL_PROMPT=0`; overriding the variable does not help. Git Credential Manager is installed
and configured (`credential.helper=manager`) but cannot open its prompt without a terminal.

**The user must run this in a normal terminal window** (fresh window so git is on PATH):
```
cd C:\Users\harry\aston-martin-financing-analysis
git push -u origin main
```
GCM will open a browser to authorise. Do **not** embed a PAT in the remote URL — it would leak into
shell history and transcripts.

After the push, verify with `git log --oneline -1` and `git status -sb` (should show `main...origin/main`).

### Original remote decision

- **Public** repository on GitHub, created manually by the user on github.com (no `gh` CLI install).
- Suggested name `aston-martin-financing-analysis`; must be created **empty** (no README/.gitignore/
  license) or the starter commit will conflict.
- ⏳ **Awaiting the repo URL**, then: `git remote add origin <url>` and `git push -u origin main`.
- GitHub attributes commits by **email**, so `harrydowney1234@gmail.com` must be registered on the
  account (Settings → Emails) for commits to link to the profile. A `@users.noreply.github.com`
  address is the alternative if the user prefers not to expose it.
- ⚠️ Note: the address in the session context was `harrydowney123@gmail.com` — **wrong**. The correct
  address is `harrydowney1234@gmail.com` (four digits). Use the corrected one.

---

## Next action

**Awaiting decision on where the repo should live** (GitHub public/private, or elsewhere), and on the
four items above. Then set the remote and push.

Reference — carried through into the final deliverables:
- The **distress-yield caveat** on the 20.16% refi rate (`SOURCES.md` §8b/B2) — must travel with the number.
- The **no-response caveat** on the downside (`SOURCES.md` §8c/C5) — real mitigation would soften it.
- The **press-usage rule** (`SOURCES.md` §9) — one labelled market-context callout each, never load-bearing.
- The **attribution robustness point**: the conclusion does not depend on the distress yield (Attribution tab).
- **Defend-list item #1** (`SOURCES.md` §8d): the conclusion is sensitive to the base case, not just the stress.
- The **alternatives table** — drafted, in `OUTPUTS.md` §11.

Phase 6 (GitHub repo) still blocked: `git` is not installed.
