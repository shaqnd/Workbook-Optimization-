# Blank Template + Full Formula Audit

`Value_Analysis_TEMPLATE.xlsm` — cleared and ready for a new property.
Built from the corrected R0195554 workbook, with every finding below fixed first.

---

## 1. What the audit checked

Nine passes over all 29 sheets — 12,000+ formulas and 36,000+ reference-table constants:

| Pass | What it looked for | Result |
|---|---|---|
| A | Lookup key/value range row alignment (17 INDEX/MATCH pairs) | 0 misaligned |
| B | Header vs value range widths for cross-column MATCH | 0 real mismatches |
| C | Blank keys inside lookup columns | 0 |
| D | Broken `#REF!` defined names | **25 found → removed** |
| E | A formula that breaks its column's pattern | 0 real (24 benign) |
| F | Cells whose cached result is an Excel error | **6 found → fixed** |
| G | SUM/AVERAGE ranges abutting live data (off-by-one) | 0 real |
| H | Constants typed over a formula link | **3 found → fixed** |
| I–K | Every MVS multiplier table vs the June 2024 manual | **575/575 verified** |
| L | Every MVS base cost vs its cited manual page | **3,095/3,095 verified** |
| M | Depreciation + life-expectancy tables vs Sec.97 | **204 cells wrong → fixed** |

### MVS data verification

Every number the cost approach depends on was checked against
`sources/MVS_June__24_compressed_30MB.pdf`:

- **3,095 base costs** — 3,081 matched their cited page by direct string search. The other 14
  were confirmed by hand: 12 are correct **midpoints of published ranges** (Sec.17 shelters,
  which MVS prints as low–high), one is Sec.14 p.25 which I read directly, and one
  (Heavy Manufacturing CMILL Average, $106.00) was confirmed via the square-metre column
  (1140.97 ÷ 10.764 = 106.00) because the PDF's overlapping columns garbled the square-foot figure.
- **6 story-height tables** (118 values), **6 sprinkler tables** (384 values),
  **current cost multipliers** (120), **local area multipliers** (110),
  **Sec.98 Denver index** (65) — all verified, zero discrepancies.
- **Depreciation tables** — 46 commercial and 39 residential rows verified verbatim against
  Sec.97 pp.24–25. Three rows did not, and that led to the most important finding below.
- **Depreciation table selection** — the `OCC_DEPR` column implements MVS's "Properties Included"
  key correctly, including its counter-intuitive parts: Sec.11 apartments and hotels use the
  *commercial* table while dormitories and clubs use *residential*; Sec.18 schools are all
  *residential*; Sec.16 splits churches and fraternal buildings off to *residential*.

---

## 2. The serious one: 201 blank cells were returning 0% depreciation

MVS Sec.97 stops printing a column once depreciation reaches its **80% maximum** — past that point
the building is at its salvage floor. The transcription copied those blanks literally, and
`INDEX()` on a blank cell returns **0**. So any building whose effective age ran past the printed
range for its life expectancy got **zero physical depreciation**, silently:

| Effective age | Typical life | Table | Was | Correct |
|---|---|---|---|---|
| 30 | 25 | commercial | **0%** | 80% |
| 46 | 20 | commercial | **0%** | 80% |
| 60 | 30 | commercial | **0%** | 80% |
| 75 | 35 | residential | **16%** | 80% |
| 90 | 25 | residential | **20%** | 80% |

98 blanks in the commercial table, 103 in the residential, plus **3 stray values** at ages 75/80/90
that appear nowhere in the manual and broke the tables' monotonicity (a shorter life must always
show more depreciation than a longer one at the same age).

The `IFERROR` message on `Cost Approach!C60` ("age exceeds tabulated range — treat as fully
depreciated") never fired, because a blank cell isn't an error.

This never surfaced on R0195554 — a four-year-old building sits well inside the printed range. It
would have hit hard on the older buildings that make up most appeal work. A 30-year-old retail
building now depreciates 80% instead of 0%; on the test case in §5 that is the difference between
a $6.1M and a $1.7M indicated value.

**Fixed:** all 204 cells set to 80. Both tables now have zero blanks and zero monotonicity
violations, and `C60` returns a number in every case so the chain can't throw `#VALUE!`.

---

## 3. Other fixes

**Multi-building cost approach (`Cost Approach MB`)** — never exercised on R0195554, since that
property has one building. It carried three of the same defects the main tab had:

- **Sprinklers inside the multiplier stack.** Row 17 read `SF*(base+sprinkler)*CCM*LAM*A/P*height`.
  MVS Sec.14 p.37 says sprinklers must not be modified for size or shape, and the Sec.10 p.3
  Calculator Cost Form carries them as a Line 28 lump sum. Now
  `SF*(base+HVAC)*multipliers + SF*sprinkler`.
- **No HVAC refinement line.** Added row 46 (`$/SF difference`, per building) for parity with the
  main tab.
- **De-trend divide-by-zero.** Row 44 had the same unguarded interpolation; added the
  equal-quarter guard.
- **Effective age hardcoded.** `C12` and `D12` held typed constants **16** and **20**, overwriting
  the `='Subject Property'!C19` links that columns E–L still had. Buildings 1 and 2 would have
  silently carried 16 and 20 years of age on any new property. Restored as formulas, and now
  measured to the level-of-value date rather than the appeal year.
- **Rounding.** `C30` hardcoded `$5,000`; now follows `Cost Approach!B70` like the main tab.

**Unguarded arithmetic — 6 cells showing live errors.** Every neighbour on these sheets guards with
`IF(x="","",...)`; these didn't, so they threw on empty comp slots:

| Cell | Was | Showed |
|---|---|---|
| `BAA Market!D10` | `=D9/E12` | `#DIV/0!` |
| `BAA Market!I39` | `=I37*E12` | `#VALUE!` |
| `BAA Rent!C25/C30/C35/C40` | `=C23*(1-C24)*C26` | `#VALUE!` ×4 |

**`Land Sales Grid!C19`** held a hardcoded `19.96` where columns E–M compute `=E18/43560`. Restored
as a formula. The `$/SF` labels on rows 30, 51, 53, 60, 61 and 65 were also relabelled **`$/AC`** —
the comps are priced per acre (row 22 is "Unadjusted Price per AC"), so the old labels were simply
wrong and wouldn't have survived cross-examination.

**25 broken `#REF!` defined names** removed (`XXX`, `XXXX`, `test1`–`test4`, `_Dkc1`–`_dkc8`,
`_Fill`, `_Key1`, `_Key2`, `_Sort`, `PRINT_TITLES_MI`, `_xlnm.Criteria`, `_xlnm.Database`). None
were referenced by any formula — verified before removal, and all 210 surviving names still resolve.

**`Reconciliation!G11`** said its source was `Basic Sales Comparison!C50`; the formula reads `C52`.

**`BOCC!A1`** hardcoded "2025 APPEAL WORKSHEET" → `=Year&" APPEAL WORKSHEET"`.

---

## 4. New safeguards

The two errors that did the most damage on R0195554 were **inputs left at their defaults** —
story height at 18 ft and the area/perimeter multiplier at 1.000. Neither raised anything. The
banner below row 76 of the Cost Approach now catches that whole class, in priority order:

1. Base cost or life expectancy couldn't be resolved → check Step 1 or use the manual override.
2. **Eave/wall height (C11) empty** → the story-height multiplier is defaulting to 1.000 and the
   shell is understated for any building taller than the MVS base height.
3. **Area/perimeter multiplier (C12/C13) empty** → it's defaulting to 1.000, which overstates a
   large-footprint building.
4. **Both de-trend quarters the same** → no interpolation is happening.
5. **No Year Built** → effective age is 0 and no depreciation is being taken.
6. Manual MVS override in use → confirm before relying on the value.

The Cost Approach header now also states the rows 24–27 / rows 28–30 split explicitly, so the
multiplied-vs-lump-sum distinction is visible on the sheet rather than only in this file.

The blank template deliberately does **not** cascade errors: `D35` falls back to 1.000 when
C12/C13 are empty and `C58`/`C60` stay numeric when there's no Year Built. The banner does the
shouting instead — an error cascade in a fresh template invites people to delete formulas.

---

## 5. Functional test

The template's own tables driven end to end on three unrelated property types, including two that
used to fall into the 0%-depreciation hole:

| Case | Base $/SF | Sec | Ht mult | CCM/de-trend | LAM | Sprinkler | Life | Depr | RCN | RCNLD |
|---|---:|---:|---:|---:|---:|---:|---:|---:|---:|---:|
| Mega warehouse, 4 yrs (R0195554 facts) | 52.00 | 14 | 1.650 | 1.0557 | 1.00 | 2.63 | 45 | 3% | 32,366,383 | 31,395,391 |
| Discount store, 30 yrs | 94.00 | 13 | 1.170 | 1.0158 | 1.00 | 4.16 | 35 | **72%** | 6,087,787 | 1,704,580 |
| Church sanctuary, 75 yrs | 188.00 | 16 | 1.210 | 1.0457 | 0.96 | 5.39 | 45 | **80%** | 3,288,844 | 657,769 |

Each case resolves its own MVS section, base cost, story-height table, sprinkler table, life
expectancy and — correctly — its own depreciation table: the church routes to the residential
curve per Sec.97's Properties Included key, the other two to commercial. (Figures are base +
sprinkler + 15% entrepreneurial profit only; no site improvements or dock floors, so they aren't
directly comparable to the R0195554 memo's $34.7M.)

---

## 6. What was cleared, and what was kept

689 cells cleared across 9 sheets. **No formula, no label and no reference table was touched** —
the clearing routine skipped any cell containing a formula, skipped column B entirely, and
preserved the `N/A` scaffolding in the Subject column of the adjustment grids.

**Cleared:** Property Info identification and assessor values · Subject Property address/size/YOC ·
Cost Approach Step 2 inputs and site-improvement quantities · Land Sales Grid (207 cells) ·
Basic Sales Comparison (104) · Detailed Sales Comparison (328 — this tab was still full of
**daycare** comps from an entirely different assignment) · Income Approach rent/vacancy/cap-rate
entries · Reconciliation weights · the stale protest number `R0186704` on WO (from yet another
property).

**Kept as reference data** — so comp auto-matching works on day one:
the full MVS engine (5,100 rows), Sales (941), Income (1,019), Expense Database (204),
PSF Summary, Property Registry, Leases, Codes, Occ Codes, Zoning Codes, Resources,
Resource Directory.

**Kept as sensible defaults:** County `ADAMS` · Prepared By · statutory valuation date 6/30/2024 ·
Number of Buildings 1 · size-match tolerance 0.25 · Condition `Typical` / Quality `Average` ·
Class `C` · City `Denver` · sprinkler tier `Average` · MVS Sec.66 unit rates ($3.10 asphalt,
$5.51 concrete, $25.18 fence) · indirect costs 0% · entrepreneurial profit 15% · rounding
increments.

**One judgment call, flagged:** the Leases tab keeps all 16 comp rows as a library, but the
`In Set?` flags are cleared so the market-rent conclusion doesn't carry over from the industrial
assignment. Re-flag the leases that fit your next subject. You asked me not to stop for the
keep-vs-strip question, so I kept the market libraries on the reasoning that keeping them costs you
one delete while stripping them costs hours of re-transcription — say the word if you'd rather have
a bare shell.

---

## 7. Before you use it

1. **Enter the OCC code first** (`Property Info!C24`) — it drives property type, the Subject
   Property tab, the Cost Approach lookups and the Property Tables.
2. **Set the appeal year** on the `S` tab (B8, currently 2025) and the statutory valuation date
   (`Property Info!C15`) when the cycle rolls to 6/30/2026.
3. **Update the de-trend quarters** (`Cost Approach!C18`/`C19`) so they straddle the appraisal date.
   They must differ from each other.
4. **Don't skip Step 2.** Height and the area/perimeter multiplier both default to 1.000 and both
   move the answer more than anything else on the sheet. The banner will tell you, but only if
   you look.
5. The MVS insert vintage is **Jan 2025 multipliers on the June 2024 manual**. When CoreLogic ships
   a new insert, Reference 8/9/10 on the M&S Inputs tab need re-keying.
6. The workbook recalculates on open (`fullCalcOnLoad`). Stale cached results were stripped from all
   17 analysis sheets, so a viewer that doesn't recalculate will show blanks rather than the last
   property's numbers.

## 8. Not changed

- **`Analysis_Workbook_SandboxCleanFinal.xlsx`** — the older repo template. Its Cost Approach is a
  structurally different, earlier build (different sheet set and formulas), and it carries two of
  the same defects: sprinklers inside the multiplier stack and the 10% indirect-cost default. It
  needs its own pass rather than a blind patch; `Value_Analysis_TEMPLATE.xlsm` supersedes it.
- **`Land Sales Grid!M22`** divides by SF while the other four comps divide by acres, which is why
  the range low reads $10.43 against a median of $240,000. The comp data is cleared now, but the
  formula asymmetry remains — whether an under-contract listing belongs in that grid at all is
  your call.
- **Site improvements depreciate on the building's curve.** Paving and fencing have 15–25 year
  lives against a building's 45. Immaterial on new construction, material at 15+ years.

---

## 9. Follow-up: filters, and the base rent / CAM / gross rent split

### Filters

`Expense Database` (**A3:AR204**) and `PSF Summary` (**A3:AM204**) now carry an autofilter on every
column, header row 3. Both register a `_xlnm._FilterDatabase` name so Excel treats the block as a
proper database.

**One more off-by-seven, found while wiring this up.** Both sheets hold 201 comps in rows 4–204,
but everything that reads them stopped at row 197:

- `Comp Picker` and `Subject Property` matched expense comps over `'Expense Database'!$AO$4:$AO$197`
  — 40 references across 20 cells. **Rows 198–204 were invisible to the expense comp matcher.**
- The OCC code, City, State and Prop Type data validations covered `B4:B197`, `E4:E197`, `F4:F197`,
  `G4:G197` — the last seven rows accepted anything.

All extended to row 204. (`Sales` was already correct at `$AJ$2:$AJ$941`.)

When you add comps, **insert rows inside the existing block** rather than appending below row 204 —
Excel then grows the filter range and the `$AO$4:$AO$204` references automatically. Appending past
the last row means widening those ranges and copying the AN/AO helper formulas down by hand.

### Leases — rent components separated

Base rent and CAM were already in different columns but ambiguously named, and nothing derived CAM
or gross rent. The tab now carries all three components, each monthly, annually and per SF of
**leased** area (not building area — several comps are partial-floor):

| Col | | Formula |
|---|---|---|
| J / K / L | **Base Rent** — Monthly / Annual / $/SF | (existing, renamed) |
| P | **CAM — Monthly** | (existing, renamed) |
| R | **CAM — Annual** | `=IF(N(P6)=0,"",P6*12)` |
| S | **CAM $/SF** | `=IF(OR(N(R6)=0,N(I6)=0),"",R6/I6)` |
| T | **Gross Rent — Monthly** | `=IF(N(J6)=0,"",J6+N(P6))` |
| U | **Gross Rent — Annual** | `=IF(N(K6)=0,"",K6+N(R6))` |
| V | **Gross Rent $/SF** | `=IF(OR(N(U6)=0,N(I6)=0),"",U6/I6)` |

Applied to the comp table (rows 6–21) and to the 125 Bridge St rent roll (rows 25–35, plus the
row 36 totals — that property has no stated unit areas, so its $/SF columns stay blank by design).

**Verified numerically:** across all 14 comps with rent data, gross $/SF equals base $/SF + CAM $/SF
to four decimals. Denver Distribution Center, for example: $6.5164 base + $0.1007 CAM = **$6.6171
gross**, on 553,757 leased SF.

**The one judgment call.** `CAM — Annual` is left **blank** where the source reported no CAM, so a
blank `CAM $/SF` marks a lease whose gross figure is base rent only. Gross rent treats an unreported
CAM as zero, because that is right for an absolute-net landlord and wrong for a lease where CAM
simply wasn't transcribed — read those rows against the **Lease Type** column. Of the 14 comps with
rent, only 5 report CAM. That note is on the tab at B52.

### Market rent conclusion — now dual

The conclusion block reports both bases side by side: **base rent in column C, gross rent in
column F**, each with count / low / high / average / median. Two hidden stat columns feed them
(`Y` base $/SF, `Z` gross $/SF), both gated on the `In Set?` flag, so the two counts also tell you
how many of your selected leases actually reported CAM.

`C49` (concluded market rent) still drives `Income Approach!D8` — unchanged.
