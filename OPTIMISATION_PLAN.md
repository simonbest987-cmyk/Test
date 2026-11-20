# Daily Tonnage Account – Semantic Model Optimisation Plan

Scope: `Daily Tonnage Account Descktop Auto Refresh by Shift Prototype upload.SemanticModel` (TMDL, compatibility level 1606).

Ground rule: **no change to business logic or reported numbers.** Every change is either
(a) proven to give an identical result from how DAX evaluates it, or (b) put off to a later
phase that needs the report definition and/or a check against real data.

## What was available, and what was not

| Item | Status |
|---|---|
| Semantic model TMDL (tables, measures, relationships, M) | ✅ Reviewed in full |
| Report definition (`.Report` folder: visuals, the fields they use, date hierarchies) | ❌ Not uploaded |
| Data (`cache.abf`) | ❌ Not opened (installing a parser was not permitted) |

Without the report I cannot see which columns, hierarchies and measure home tables the visuals use.
Removing, renaming, moving or retyping any model object could break visuals, so those changes are in
**Phase 2**. Phase 1 changes only *how* existing objects are calculated, never *what* they are.

---

## 1. Model overview

* **Source:** a single Excel workbook on SharePoint
  (`.../upload data/\daily_tonnage_account by Shift Prototype 2026 v2.xlsx`), loaded by **16 separate
  queries**, each calling `Excel.Workbook(Web.Contents(<hard-coded URL>))`.
* **Imported tables (15):** T_Tracking_Account (large transaction table, about 90 source columns + 40 calculated
  columns), T_Delays, T_Rolling_Times, Shift_Dates, Calendar, Unique_Week, Operating_Costs,
  Product_Costs, Extra_Processing, Sorting_Bank, External_Scrap, T_Hotsaw_Scrap, T_Annual_Plan_Matrix,
  Gas_Charges, Electric_Charges, T_Summary_Utility_Charges.
* **Calculated tables (7):** Summarised Delays, Summarised Rolling, Summarised Cuts Weight,
  Summarised Cuts Weight Total, Summarised Blm Weight, Summarised Prime Weight, Summarised Defects.
* **Auto date/time:** on. **13 hidden LocalDateTables** + 1 template, one for each datetime column.
* **Relationships:** 50. Of these, 13 are to auto-date tables, **19 are inactive**, and **5 are bidirectional and one-to-one**.

## 2. Findings

### 2.1 Calculated columns that call measures (context transition on every row) – refresh cost  ★ High

When a calculated column calls a measure, every row of the table becomes a filter on *every column of that
table* (context transition). On T_Tracking_Account (130 columns) this is the most expensive pattern in the model.

| Column | Measure called | Analysis |
|---|---|---|
| `T_Tracking_Account[36 Speed Total £]` | `[SOW Rate Total]` | Operating_Costs has **no relationships**, so the row filter can never reach it. The measure returns the same constant on every row. |
| `Extra_Processing[Standard Worked]` | 5 × `[... Manned Hours]` | Same: Operating_Costs is unrelated, so each branch is a constant per cost centre. |
| `Extra_Processing[Standard Rate]` | 5 × `[... Process Rate]` | Same. |
| `Calendar[Summarised Rolling]` | `[Rolling Time Related]` | Calendar only filters *towards* itself (it is the many side, single direction), so the row filter never reaches Summarised Rolling. The result is the grand total. |
| `'Summarised Cuts Weight'[Fin Scrap t]` | `[Prime Output]` | The table is a SUMMARIZE output, so rows are unique. Context transition selects exactly the current row, so `[Prime Output]` = `[H/Saw Weight] + [Defects]` for that row. |

**Action (Phase 1):** rewrite each one with row-level or `SUMX(FILTER(Operating_Costs …))` expressions
that give the same values without context transition.

### 2.2 Inefficient DAX measures – query cost

| Measure | Issue | Action |
|---|---|---|
| `Calendar[Pro Rata Cuts Wght1]` | `FILTER(ALL('Summarised Cuts Weight'), key = MAX(key))` loops over the whole table and evaluates MAX once per row. It sits under 6 or more dependent measures (PR Rolling, Worked Time1, Speed (A), Avail (A) …). | Phase 1: capture MAX once in a variable, then `REMOVEFILTERS` + column predicate. Same semantics (ALL(table) as a filter = REMOVEFILTERS on the expanded table). |
| `Summarised Rolling[Rolling Time Related]`, `Unique_Week[Worked Time]`, `Summarised Blm Weight[42 Speed Stnd]` | `SUMX(RELATEDTABLE(T), T[c])` inside a measure. A measure has no row context, so this is just `SUM(T[c])` with an unnecessary table iteration. | Phase 1: replace with `SUM`. |
| ~35 measures `CALCULATE(SUM(Operating_Costs[x]), [Cost Centre] IN {n})` | Efficient already. Repetitive, but not a performance problem. | None (a calculation group or a measure table is cosmetic, and moving measures breaks visuals). |

### 2.3 Calculated tables and the "Summarised" mesh (reconciliation tables) – ★ High, structural

* Five calculated tables re-aggregate T_Tracking_Account at **Week × Product** grain, each with a
  concatenated **string** key `[Week]&"_"&[Product Code]&"_"&[product]`, joined to each other by
  **bidirectional one-to-one** relationships (Prime↔Blm, Cuts↔Defects, Delays↔Rolling), plus 15 inactive
  relationships that are never activated with USERELATIONSHIP.
* `Summarised Cuts Weight` then uses LOOKUPVALUE to pull Blm, H/Saw (Prime), Defects, Delays and Rolling back in.
  It is already the de-facto Week×Product fact table.
* T_Tracking_Account calculated columns look up into calculated tables that are built *from*
  T_Tracking_Account (`Lookup Shift` → `Summarised Cuts Weight Total` → `Summarised Cuts Total`). This creates a long chain of dependent calculations at refresh.

**Correctness observations (NOT changed, for the business to confirm):**
1. `T_Tracking_Account[Summarised Manned Total]` returns `'Summarised Rolling'[Summarised Rolling]`
   but searches on `'Summarised Delays'[WEEK_ABB Product]`. It therefore returns blank for any product-week
   that has rolling time but **no delays**, which flows into `PR Manned Total` and `Availability Actual Measure`.
2. `Summarised Cuts Weight` and `Summarised Rolling` reach `Unique_Week` **only through Summarised Delays**
   (their own week relationships are inactive). When filtering by week, product-weeks with **zero delays drop out** of
   Cuts and Rolling totals.
3. `'Summarised Cuts Weight'[Delays%] = .35` and `[Availability Standard] = 1-.40` are hard-coded constants.

**Action (Phase 2):** replace the mesh with one Power Query–built fact `Fact_WeekProduct` (Week, Product key,
Cuts, Bloom, Prime, Defects, Delays, Rolling) with integer keys to `Dim_Week` and `Dim_Product`, and single-direction
relationships. This needs the report (visuals use columns from these tables) and before/after numbers,
especially because observations 1 and 2 mean a "pure" star schema would *change* totals unless the current
behaviour is deliberately reproduced.

### 2.4 Large transaction table – T_Tracking_Account  ★ High

* **Denormalised product cost card:** about 45 columns (`42_mill_*`, `36_mill_*`, `cathode_*`, `darlo_*`, `cold_saw_*`,
  `restraight_*`, `crane_rail_*`, `Wind*`, `product_cost_total` …) repeat the `Product_Costs` row on every
  transaction. → Phase 2: `Dim_Product` (from Product_Costs), keep only the key on the fact.
* **36 pure row-arithmetic calculated columns** (`Total Output`, `*OSP`, `42/36 Scrap £`, `Scale £`, `Roll Tariff £`,
  `Production Value £`, tonnage columns …). These would be cheaper as Power Query columns, or better still computed in the
  source workbook, which already computes the product-cost joins. Not moved yet: DAX BLANK arithmetic
  (`BLANK + x = x`) differs from M null arithmetic (`null + x = null`), and columns such as `Column1` (defects) are
  likely to be sparse. Moving them safely needs an explicit null-handling rule for each column **and** a data comparison.
  → Phase 2, with the validation queries in §4.
* **High-cardinality columns that are probably not used in reporting:** `Time at Saw1–4` (text timestamps), `Cuts1–4`,
  `Piece1–4`, `Cast`, `Sequence`, `Strand`, `Cut`, `Order`, `Item`, `Customer`, `Column2`, `Column3`,
  `Non prime`. Several are typed `any` (they load as text). → Phase 2: remove the unused ones once the report has been checked.
* Five meaningless hierarchies on numeric £/tonnage columns (e.g. `36 Mill Scrap £ Hierarchy`, twice). → Phase 2: remove if unused.

### 2.5 Financial calendar logic  ★ Medium

* **Three overlapping calendar tables:** `Shift_Dates` (table), `Calendar` (*sheet* load of the same data, using
  PromoteHeaders), and `Unique_Week` (one column). The weekly key appears as `Week` (double, shift-fractional),
  `Week2`, `Week ABB`, `Week No`, `ABB Week`, and `Column1` in Sorting_Bank.
* `Calendar` loads the whole sheet: 29 junk columns (`Column10`–`Column40`, `KPI`, `Yearly Budget`,
  `Weekly Budget`, `Week_1`) from an adjacent block on the sheet, all as text.
* Calendar is used as the **home table for about 100 measures**. That is fine for performance, but it makes Calendar hard to replace.
* **Relationships across data types:** `Extra_Processing[Week ABB]` (text), `T_Hotsaw_Scrap[Week No]` (text) and
  `Sorting_Bank[Column1]` (text) all join to `Unique_Week[Week No]` (whole number).
* **13 auto date/time tables**, including ones on columns that are not meaningful dates for
  slicing (`T_Delays[End]`, `T_Rolling_Times[End]`, `Gas_Charges[Start/End Date]`, two `Time at Saw1 End` columns …).

**Action (Phase 2, needs report):** one `Dim_Date` (date + shift + financial week + week number) built in Power Query
from Shift_Dates; `Dim_Week` from it; whole-number week keys everywhere; turn off auto date/time; mark `Dim_Date` as a date table;
trim Calendar to the columns it needs, or retire it once its measures have moved to a measure table.

### 2.6 SAP / utility reconciliation  ★ Low (small tables)

* `T_Summary_Utility_Charges` holds the SAP-vs-invoice reconciliation (`Gas SAP`, `Elec SAP`, `* Difference`),
  calculated **upstream in Excel**. The model has no NAV source and no SAP/NAV matching logic, so there is
  nothing to optimise in DAX here.
* `Electric_Charges ↔ T_Summary_Utility_Charges` is an **auto-detected, bidirectional one-to-one** relationship on
  the text key `Area_Energy Bill`, and `Gas_Charges → T_Summary_Utility_Charges` is many-to-one on the same key. → Phase 2:
  an `Area_Bill` dimension with single-direction relationships (needs a report check because of the bidirectional filtering).

### 2.7 Power Query / source

* The same hard-coded URL is used 16 times. → **Phase 1:** a `SourceWorkbookUrl` parameter + a shared `DTA_Workbook` query.
  This mainly helps maintenance (switching between test and production, one credential), and gives the mashup engine one shared source.
  The bigger refresh win is moving the source off a single large workbook (Dataflow / SharePoint list / SQL) → Phase 3.
* `Errors in Operating_Costs` and `Errors in Operating_Costs (2)` are leftover, unloaded error-report queries
  from 2023, with two query groups. → **Phase 1:** remove them (they are not loaded and nothing references them).

---

## 3. Phased plan

### Phase 1 – implemented now (results identical by construction)

| # | Change | File(s) | Benefit |
|---|---|---|---|
| 1 | `36 Speed Total £`: remove context transition (inline `SOW Rate Total` as `SUMX(FILTER(Operating_Costs, CC=303605))`) | T_Tracking_Account.tmdl | Refresh: no 130-column context transition per transaction row |
| 2 | `Standard Worked` / `Standard Rate`: SWITCH on Area → cost centre, one `SUMX(FILTER())`, no measure calls | Extra_Processing.tmdl | Refresh |
| 3 | `Calendar[Summarised Rolling]` = `SUM('Summarised Rolling'[Summarised Rolling])` | Calendar.tmdl | Refresh |
| 4 | `Fin Scrap t`: replace `[Prime Output]` with the row-level `[H/Saw Weight] + [Defects]` | Summarised Cuts Weight.tmdl | Refresh |
| 5 | `Pro Rata Cuts Wght1`: variable + REMOVEFILTERS instead of FILTER(ALL()) | Calendar.tmdl | Query time, for the Speed / Avail / Worked Time measures that depend on it |
| 6 | `Rolling Time Related`, `Worked Time`, `42 Speed Stnd`: SUMX(RELATEDTABLE()) → SUM | 3 files | Query time (minor), readability |
| 7 | URL parameter + shared workbook query, used by all 16 partitions | expressions.tmdl + 16 table files | Maintainability, one source |
| 8 | Delete orphan `Errors in Operating_Costs` queries and their groups | expressions.tmdl, model.tmdl | Cleanup |

No table, column, measure, hierarchy or relationship is renamed, removed or retyped, so the report keeps working.

### Phase 2 – needs the `.Report` folder + a before/after data comparison

1. Turn off auto date/time; build `Dim_Date` / `Dim_Week` (Power Query) with whole-number week keys; fix the 3 relationships across data types.
2. Remove the unused high-cardinality and junk columns (T_Tracking_Account list in §2.4, Calendar `Column10–40`, `Column1–3` in
   T_Delays/T_Rolling_Times, `Column1–5` in Electric_Charges, `Column1/2` in Product_Costs); split datetime columns into date + time where only the date is used.
3. `Dim_Product` from Product_Costs; drop the 45 denormalised cost columns from T_Tracking_Account. Calculated columns that use them become
   `RELATED()`, or move to Power Query.
4. Move the 36 row-arithmetic calculated columns on T_Tracking_Account into Power Query, with explicit BLANK-equivalent null handling,
   or turn the ones that are only ever summed into `SUMX` measures.
5. Replace the five Summarised tables and their bidirectional one-to-one mesh with one `Fact_WeekProduct` + integer keys, keeping today's behaviour
   (observations 1–2 in §2.3) unless the business decides to correct it.
6. Measure table (`_Measures`), display folders; remove accidental hierarchies; set `summarizeBy: none` on keys and rates.
7. Make the Electric/Gas ↔ Summary relationships single-direction through an `Area_Bill` dimension.

### Phase 3 – source

Move the heavy shaping (product cost card join, Week×Product aggregation, SAP reconciliation) into the source (SQL view / Dataflow),
and keep the Excel workbook only for manual inputs (Annual Plan, Operating Costs).

## 4. How to validate Phase 1 (DAX Studio / DAX query view)

Run each query against the **original** model and the **updated** model and compare the results, or run the
"old vs new" queries below against the updated model; each should return **no rows**.

```dax
// 1. 36 Speed Total £ – old expression recomputed via the measure
EVALUATE
FILTER (
    ADDCOLUMNS ( T_Tracking_Account, "@old",
        -(T_Tracking_Account[PR Worked Time Total] - T_Tracking_Account[PR Worked Time Total Stnd]) * [SOW Rate Total] ),
    ABS ( [@old] - T_Tracking_Account[36 Speed Total £] ) > 1e-9
)

// 2. Extra_Processing Standard Worked / Standard Rate
EVALUATE
FILTER (
    ADDCOLUMNS ( Extra_Processing,
        "@oldWorked", SWITCH ( Extra_Processing[Area],
            "Cold Saw", [Cold Saw Process Manned Hours], "Crane Rail", [Crane Rail Process Manned Hours],
            "Cathodes", [Cathode Process Manned Hours], "Darlington", [Darlington Manned Hours],
            "Sidepress", [Sidepress Process Manned Hours] ),
        "@oldRate", SWITCH ( Extra_Processing[Area],
            "Cold Saw", [Cold Saw Process Rate], "Crane Rail", [Crane Rail Process Rate],
            "Cathodes", [Cathode Process Rate], "Darlington", [Darlington Process Rate],
            "Sidepress", [Sidepress Process Rate] ) ),
    [@oldWorked] <> Extra_Processing[Standard Worked] || [@oldRate] <> Extra_Processing[Standard Rate]
)

// 4. Fin Scrap t
EVALUATE
FILTER (
    ADDCOLUMNS ( 'Summarised Cuts Weight', "@old",
        'Summarised Cuts Weight'[H/Saw Weight] - ( [Prime Output] * 'Summarised Cuts Weight'[fin_yield] ) ),
    ABS ( [@old] - 'Summarised Cuts Weight'[Fin Scrap t] ) > 1e-9
)

// 5/6. Headline measures by week – compare before vs after
EVALUATE
SUMMARIZECOLUMNS (
    Unique_Week[Week No],
    "Speed A", [Speed (A)], "Avail A", [Avail (A)], "Worked Time", [Worked Time], "Avail", [Avail],
    "42 Speed Stnd", [42 Speed Stnd], "Total Variance", [Total Vairance Calc], "Difference", [Difference],
    "36 Variance", [36 Variance £], "42 Variance", [42 variance £]
)
```

## 5. Data protection note

The `Test` repository is **public**. The uploaded zip contains `.pbi/cache.abf`, which holds a full copy of the
imported data, and the TMDL exposes the internal SharePoint site URL. Consider making the repo private and removing the
zip from its history. The `.gitignore` added here keeps `cache.abf` and `localSettings.json` out of future commits.
