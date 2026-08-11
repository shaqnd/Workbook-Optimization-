# R0195554 — Cost Approach Reconciliation

**Subject:** 22600 Interstate 76, Brighton CO · Adams County account R0195554 · parcel 1569014‑01‑001
**Improvement:** 352,240 SF Class C "Mega (Storage/Distribution) Warehouse" (OCC 584), built 2020
**Tax year:** 2025 · level of value 6/30/2024
**Documents compared:** Adams County *Cost Breakdown Sheet* (report date 8/11/2026) vs. the Cost Approach tab of `Value_Analysis_R0195554.xlsm`
**Cost authority:** Marshall Valuation Service, June 2024 edition (`sources/MVS_June__24_compressed_30MB.pdf`)

---

## 1. Short answer

The two never lined up for eleven separate reasons, not one. **Eight were errors in our workbook** — two of
them structural formula bugs that would have recurred on every future property, and one of which was
failing silently. **One is an error on the assessor's card.** The remaining **two are legitimate
methodological differences** between a fee appraisal and a mass-appraisal CAMA run, and those should not be
reconciled away.

Before the fix the workbook indicated **$33,880,000** total (improvements $28,880,724 + land $5,000,000)
against the assessor's **$36,924,394 for improvements alone** — a gap of roughly $8.0M on the improvements,
large enough that neither number could be defended in a hearing.

After the fix the workbook indicates **$38,630,000** total, of which **$33,627,038** is depreciated
improvements. That is **$3,297,356 (8.9%) below the assessor's cost card**, and every dollar of the
remaining difference is now traceable to a named line item (§4).

The concluded value on the Reconciliation tab is **unchanged at $35,225,000** — the cost approach carries
0% weight there. What changed is that the cost approach is now defensible if anyone asks about it.

---

## 2. What was wrong in the workbook

### 2.1 Story height was 18 ft — the biggest single error

`Cost Approach!C11` held **18 ft**. The Adams County record card reads **Story Ht: 42**. MVS Sec.14 p.39
keys story-height multipliers to a 14 ft base:

| Height | Multiplier |
|---|---|
| 18 ft (what the workbook used) | **1.086** |
| 42 ft → 40 ft row (assessor, and now us) | **1.650** |

18 ft is a 1980s tuck-under warehouse. A 2020 mega distribution building has 36–40 ft clear. This one
input understated the building shell by 52%. **Fixed → 42.** Worth confirming at inspection, because
nothing else in the calculation moves the number this much.

### 2.2 The Floor Area–Perimeter multiplier was never entered

`C12`/`C13` were both left at the template default of **1.000**. MVS Sec.14 p.39, entered at 352,240 SF
and a 2,644 ft average perimeter, two-way interpolates to **0.853**:

```
350,000 SF row:  .853 @ 2,600 ft   .857 @ 3,000 ft   → .8534 @ 2,644 ft
400,000 SF row:  .848 @ 2,600 ft   .853 @ 3,000 ft   → .8486 @ 2,644 ft
interpolate to 352,240 SF                            → 0.853
```

This is exactly the 0.8530 the assessor used, independently derived. Leaving it at 1.000 overstated the
shell by 17%. **Fixed → 0.853 in both cells.**

Note these two errors ran in opposite directions and partly masked each other, which is probably why
neither was caught: 1.000 × 1.086 = 1.086 against the correct 0.853 × 1.650 = 1.408.

### 2.3 A phantom second building

Row 26 carried a hard-coded **"Building 2" — 2,400 SF @ $42.75 = $102,600**, while Subject Property!C8
says one building. The $42.75 is the MVS Sec.14 p.37 average wood-frame canopy rate, so this was left over
from an earlier property. **Removed** (row 26 is now a labelled spare refinement line).

### 2.4 Sprinklers were being run through the size and shape multipliers

`F31` summed rows 24–30 — base cost *and* sprinklers — and `G38` then multiplied the whole thing by the
product of the multipliers. Two MVS passages say that is wrong:

* Sec.14 p.37: *"Sprinklers should not be modified for size or shape."*
* Sec.10 p.3, the official Calculator Cost Form: sprinklers are a **Line 28 lump sum**, added *after*
  Lines 18–21 (stories / story height / floor-area-perimeter) and Lines 23–24 (current cost / local).

With the corrected multipliers this bug alone would have inflated sprinklers by 48.6%. The assessor did
not make this mistake — their card carries sprinklers as a flat 352,240 × $2.68.

**Fixed.** Rows 24–27 are now the multiplier-adjusted block, rows 28–30 are lump sums, and
`G38 = F37*F31 + SUM(F28:F30)`. The section labels say so on the sheet.

### 2.5 The de-trend to the valuation date was silently switched off

`C18` and `C19` (the MVS Sec.98 Denver index bracket quarters) were **both set to "Apr 2024."** The
interpolation in `D33` then divides by `DATEVALUE("1 Apr 2024") − DATEVALUE("1 Apr 2024")` = 0, the
`IFERROR` swallowed the `#DIV/0!`, and the de-trend factor quietly fell back to **1.000**. The displayed
multiplier of 1.060 looked plausible, so nothing flagged it.

**Fixed two ways:** `C19` set to **Jul 2024** so the quarters straddle 6/30/2024, and `D33` rewritten with
an explicit `IF($C$18=$C$19, ...)` guard so the same-quarter case returns that quarter's index instead of
failing into 1.000.

```
Apr 2024 Class C = 1.012 · Jul 2024 Class C = 1.004 · 90/91 of the way = 1.0041
D33 = CCM 1.060 / 1.0041 = 1.0557   (was 1.060)
```

### 2.6 A 10% indirect-cost allowance that MVS already includes

`C48` added **10% for "financing, permits, professional fees, insurance during construction"** — $2,378,189.
Every one of those is already inside the MVS calculator cost. MVS Sec.1 p.3, *What the Costs Contain*:

> (1) …average architects' and engineers' fees. These, in turn, include plans, plan check and nominal
> building permits, and surveying… (3) Normal interest on only the actual building funds during period of
> construction and processing fee or service charge is included… (7) Contractors' overhead and profit
> including job supervision, workmen's compensation, fire and liability insurance… are included.

**Set to 0%**, with the tooltip rewritten to say what the line *is* legitimately for — the Sec.1 p.3
*What They Do Not Contain* list: impact and tap fees, off-site work, entitlement, feasibility, land carry.

### 2.7 Depreciation was struck one year past the valuation date

`C58` read `Subj_Age`, which is `appeal year (2025) − YOC (2020) = 5`. At the 6/30/2024 level of value the
building is **4**. On the MVS Sec.97 p.24 commercial curve at a 45-year life that is the difference between
4% and 3%.

**Fixed → `=IFERROR(YEAR($C$17)-Subj_YOC,Subj_Age)`**, which measures age to the appraisal date in `C17`.
Gives 4 years, matching the assessor's own effective age. Still overtypeable if inspection supports a
judgment figure.

### 2.8 Site improvements were another property's numbers

Row 42 carried **275,000 SF of concrete paving** and nothing else. The site is 869,458 SF with a 352,240 SF
footprint, so 275,000 SF of concrete and zero asphalt was never plausible. Replaced with the assessor's
measured quantities priced at MVS Sec.66 rates:

| Line | Quantity | MVS rate | Total | (assessor's rate) |
|---|---|---|---|---|
| Asphalt paving | 316,000 SF | $3.10 /SF — Sec.66 p.1, 4″ asphaltic concrete | $979,600 | $1.55 |
| Concrete paving | 71,000 SF | $5.51 /SF — Sec.66 p.1, 6″ concrete | $391,210 | $3.07 |
| Chain link fence, 8 ft | 1,500 LF | $25.18 /LF — Sec.66 p.4, #11 wire at $26.50 less the 5% large-installation tier | $37,770 | $12.50 |
| **Total site improvements** | | | **$1,408,580** | *$726,520* |

Dock-height floors (296,000 SF @ $2.00 = $592,000, MVS Sec.14 p.27) were added as a **lump sum** on row 29.
MVS says to add them to the base cost, which would push them through the multipliers; they are carried
unmultiplied here because the area/perimeter and story-height factors describe wall economics, not slab
thickness — and because that is how the assessor treated them, so it is not a point of dispute.

Our yard rates run roughly double the assessor's across the board. That is a defensible MVS-vs-CAMA
difference, and it works *against* us, so it is worth knowing before a hearing rather than during one.

### 2.9 One error on the assessor's side: gross HVAC instead of net

The card adds **$14.85/SF for a "Package Unit" — $5,230,764**, and that is the largest single number in
their whole calculation. Two problems with it:

1. **It is gross, not the difference.** MVS Sec.14 p.36: *"If the heating found in the building being
   appraised is different from that indicated for the base being used, take the difference between the
   costs of the two and add to or subtract from the base square foot cost."* The Sec.14 p.25 Mega Warehouse
   base cost is published **with space heaters** ($3.10/SF, moderate climate). Even granting package A.C.
   ($13.05/SF), the correct add is **$13.05 − $3.10 = $9.95/SF**, not $14.85. Their own method overstates
   by roughly $1.7M.
2. **A 352,240 SF distribution warehouse is not fully air-conditioned.** MVS publishes this occupancy with
   space heaters precisely because that is what these buildings have; conditioned office areas are already
   inside the Average quality tier.

The workbook now has an explicit **Heating / Cooling / Ventilation Refinement** line (row 25) that takes a
*difference*, set to **$0.00** — the building is space-heated as the base cost assumes. If inspection shows
the warehouse floor really is on package A.C., enter **$9.95**, not $14.85.

*Aside:* the assessor also adds HVAC **after** the size/height multipliers. The MVS form adds it at Line 14,
**before** them. Applying the multipliers would have made their number larger still, so this ordering error
happens to run in our favor.

---

## 3. The two legitimate differences

These are not errors on either side. They are judgment calls, and they are most of what is left.

**Entrepreneurial profit — 15%, $4,521,490.** MVS Sec.1 p.3 explicitly excludes it: *"Costs of land planning
or preliminary concept and layout for large developments inclusive of entrepreneurial incentives or
developer's overhead and profit are not included."* Including it is correct appraisal practice for market
value; the assessor's mass-appraisal card carries none. This is the single largest reason our RCN can
exceed theirs, and 15% is at the upper end of the typical 10–20% range — **worth a second look before you
rely on it.**

**The assessor's +10% "Design Adjustment," $3,483,416.** A CAMA factor with no MVS analog and no stated
basis on the card. There is nothing to replicate in the workbook. If the cost approach ever becomes
contested, this is the line to make them justify.

---

## 4. The bridge

Both columns are improvements only, depreciated. Land is excluded from both.

| | Assessor | Workbook (corrected) |
|---|---:|---:|
| Base $/SF — MVS Sec.14 p.25, Class C / Average | 52.00 | 52.00 |
| Cost × local multipliers | 1.0500 × 1.0100 = 1.0605 | 1.0557 × 1.000 |
| Perimeter × stories × story height | 0.8530 × 1.0000 × 1.6500 | 0.853 × 1.000 × 1.650 |
| Shell $/SF | 77.62 | 77.27 |
| HVAC add | +14.85 | 0.00 |
| **Adjusted base $/SF** | **92.47** | **77.27** |
| Building RCN (adjusted base × 352,240 SF) | 32,571,633 | 27,218,289 |
| Sprinklers | 944,003 @ $2.68 | 926,391 @ $2.63 |
| Dock-height floors | 592,000 | 592,000 |
| Other yard improvements | 726,520 | 1,408,580 |
| **Replacement cost new (before soft costs)** | **34,834,156** | **30,145,260** |
| Design adjustment +10% | 3,483,416 | — |
| Indirect costs | — (in MVS) | 0 — (in MVS) |
| Entrepreneurial profit +15% | — | 4,521,789 |
| **Total RCN** | **38,317,572** | **34,667,050** |
| Physical depreciation | (1,393,178) @ 4%, eff. age 4 | (1,040,011) @ 3%, eff. age 4 / life 45 |
| **Depreciated improvements** | **36,924,394** | **33,627,038** |
| per SF of building | $104.83 | $95.47 |

**Reconciliation of the $3,297,356 gap:**

| | $ |
|---|---:|
| Assessor's gross HVAC add, which we carry at zero | (5,230,764) |
| Assessor's +10% design adjustment, which has no MVS analog | (3,483,416) |
| Our entrepreneurial profit at 15%, which they carry at zero | +4,521,789 |
| Our yard improvements at MVS rates vs. their roughly half-MVS rates | +682,060 |
| Cost/local multiplier difference on the shell (1.0605 vs. 1.0557) | (122,579) |
| Sprinklers, $2.68 vs. $2.63 | (17,612) |
| Less depreciation: 3% of a smaller base vs. 4% of a larger one | +353,167 |
| **Net** | **(3,297,356)** |

---

## 5. Cross-checks

* **Adjusted base cost.** Their $92.47/SF and our $77.27/SF differ by $15.20, of which $14.85 is the HVAC
  add. Strip that out and the two independent builds agree to within **0.5%** — good evidence that the
  occupancy, class, quality, base rate and refinement multipliers are all now right on both sides.
* **RCN per SF.** $98.42 including profit and site work, against roughly $85–115/SF for Denver Class-A
  distribution construction at the 2024 level of value. In range.
* **Sanity.** The property sold 1/14/2022 for $49,900,000 ($141.66/SF). A $38.63M cost indication including
  land sits well below it, as it should for a cost approach on a four-year-old building.
* **Their own file disagrees with itself.** The card concludes $36,924,394 for improvements, but the NOV
  carries improvements at **$32,219,767** — $4.7M lower. The cost card is evidently *not* what set the
  assessed value, which is worth knowing if they try to lead with it.

---

## 6. What changed in the file

Everything below is in `Value_Analysis_R0195554.xlsm`, tab **Cost Approach**, except where noted. VBA,
form controls, drawings and cell comments are byte-identical to the original — only the four sheet XML
parts that needed to change were touched. The workbook is flagged to recalculate on open.

| Cell | Was | Now |
|---|---|---|
| `C11` | 18 | **42** — story height per the county card |
| `C12`, `C13` | 1.000 | **0.853** — MVS Sec.14 p.39 interpolated |
| `C19` | Apr 2024 | **Jul 2024** — brackets now straddle 6/30/2024 |
| `D33` | de-trend collapsed to 1.000 on equal quarters | explicit `IF($C$18=$C$19,…)` guard; now **1.0557** |
| `B25:F25` | Sprinklers | **Heating/Cooling/Ventilation Refinement ($/SF difference)**, default 0.00 |
| `B26:F26` | Building 2 — 2,400 SF @ $42.75 | cleared, labelled spare refinement line |
| `B28:F28` | empty | **Sprinklers** as a lump sum (formula moved from row 25 unchanged) |
| `B29:F29` | empty | **Dock-Height Floors** — 296,000 SF @ $2.00 |
| `F31` | `SUM(F24:F30)` | `SUM(F24:F27)` — multiplier-adjusted items only |
| `G38` | `F37*F31` | `F37*F31+SUM(F28:F30)` — lump sums added after the multipliers |
| `C40` | 0 | **316,000** SF asphalt |
| `C42` | 275,000 | **71,000** SF concrete |
| `B43:D43` | "Other1", empty | **Chain Link Fence, 8 ft** — 1,500 LF @ $25.18 |
| `C48` | 10% | **0%** — already inside the MVS cost |
| `C58` | `=Subj_Age` (appeal year) | `=IFERROR(YEAR($C$17)-Subj_YOC,Subj_Age)` (valuation date) |
| `Property Info!C15` | 6/30/2025 | **6/30/2024** — the 2025–26 cycle level of value, now agreeing with `Cost Approach!C17` |

Every changed line carries a rewritten note in column I with its MVS section and page cite, so the
reasoning travels with the file.

---

## 7. Left alone — your call

1. **Entrepreneurial profit at 15%** (`C50`). Judgment, and the largest remaining lever. §3.
2. **Story height of 42 ft.** Taken from the assessor's card, not from our own measurement. Verify at
   inspection — it moves the number more than anything else on the sheet.
3. **Dock-height floor area of 296,000 SF.** Also the assessor's figure, unverified.
4. **Landscaping** (`C45`), left at zero to match the assessor. Roughly 130,000 SF of the site is outside
   the building and the paved areas; if detention, landscaping or yard lighting were built, they belong here.
5. **Land Sales Grid — units are mislabelled.** The comps are priced **per acre** (row 22, "Unadjusted Price
   per AC"), but rows 51 and 53–57 and the conclusion block at rows 60/61/65 all read "$/SF." `C61` is
   labelled "Land Area (SF)" and pulls **19.96 acres**. The math is internally consistent and the $5,000,000
   conclusion is right, but the labels would not survive cross-examination. Separately, comp 5's `M22`
   divides by SF (`M18`) while the other four divide by acres, which is why the range low reads **$10.43**
   against a median of $240,000 — that comp is corrupting the Low/Average/Median stats. Not touched, since
   whether an under-contract listing belongs in the grid at all is your call.
6. **Site improvements depreciate on the building's curve.** Paving and fencing have 15–25 year lives
   against the building's 45. Immaterial at 4 years old; it will not be at 15.
7. **The blank template** (`Analysis_Workbook_SandboxCleanFinal.xlsx`) carries two of the same structural
   issues — sprinklers inside the multiplier stack (`F31`/`G38`) and the 10% indirect default in `C48` — so
   they will recur on the next property. Its Cost Approach tab is an older, structurally different build,
   so it was left untouched rather than patched blind. Worth a separate pass.
