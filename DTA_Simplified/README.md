# DTA_Simplified – Variance Bridge model and report

A separate, cleaned-up Power BI project that shows the **variance bridge**: the eight variance components
that make up `Total Vairance Calc`, compared with the 42 + 36 mill accounts (`Total Variance Calc1`).
The original Daily Tonnage Account file is not changed.

**Numbers are identical to the original model.** Every measure, calculated column and calculated table used by the
bridge is copied unchanged (the Phase 1 versions in the main model). Only unused tables, columns,
relationships and auto date tables have been removed.

## Open it

1. Download the repo branch (GitHub → **Code → Download ZIP**) and unzip it to a **short path**, e.g. `C:\PBI\`.
   Long paths (Downloads, OneDrive) can break the unzip on Windows.
2. Open the `DTA_Simplified` folder and double-click **`DTA_Simplified.pbip`**. Open this one, not the `.pbip` at the top of the repo.
3. **Home → Refresh.** Sign in to SharePoint with **Organizational account** if asked.
   The source file is the same workbook as the original. Change it under **Transform data → Edit parameters → `SourceWorkbookUrl`**.

## What's in the report

| Page | Contents |
|---|---|
| **Variance Bridge** | Week slicer, 4 cards (Bridge total, Mill Accounts total, Unexplained Difference, Reconciliation Status), a waterfall of the 8 steps, and a step table |
| **Weekly Reconciliation** | One row per week: each bridge step, Bridge total, 42 / 36 mill variance, Mill Accounts total, Difference, Status |

**Reconciliation Status** shows "Balanced" when the bridge and the mill accounts agree to within £1. Otherwise it shows
"Out of balance by £x". This shows the gap that exists in the current logic; it does not force it to zero.

## What's in the model

| Table | Role | Visible |
|---|---|---|
| `Dim_Week` | Week number (from the workbook's `Unique_Week` table). **The only slicer table.** | ✅ |
| `_Measures` | Folder **Variance Bridge**: 15 clearly named measures. Folder **Calculation steps (original logic)**: the 42 original measures, unchanged and hidden | ✅ |
| `Variance Bridge Steps` | The 8 step labels for the waterfall | ✅ |
| T_Tracking_Account, T_Delays, Extra_Processing, Operating_Costs, Product_Costs, Shift_Dates, T_Rolling_Times, 6 × Summarised … | Unchanged calculations, hidden | hidden |

Compared with the original model:

| | Original | Simplified |
|---|---|---|
| Tables | 38 (incl. 15 auto date) | 16 |
| T_Tracking_Account columns | 132 | 57 |
| Relationships | 49 (17 inactive, 14 auto date) | 7 (none inactive) |
| Auto date/time | On | Off |

### Why the numbers stay the same

* Each Power Query query is the original one, plus a final `Table.SelectColumns` step. Removing columns never changes row counts or values.
* Week filters reach each table **along the same paths as before**:
  `Dim_Week` → T_Tracking_Account, T_Delays, Summarised Delays, Extra_Processing; Summarised Delays → Summarised Cuts Weight;
  Summarised Delays ↔ Summarised Rolling and Summarised Cuts Weight ↔ Summarised Defects (the same one-to-one relationships).
  The relationships that were removed never carried a filter into any bridge table.
* Operating_Costs is still unrelated, so the rates are the same constants as before.

**Caveat:** the numbers are identical when you filter by **week number** (Unique_Week in the original, `Dim_Week` here).
Do not add product or date slicers from the hidden tables. They would filter differently from the original report.

## Check it

Run this in **DAX query view** in both files and compare. In the original, replace `Dim_Week` with `Unique_Week`.

```dax
EVALUATE
SUMMARIZECOLUMNS (
    Dim_Week[Week No],
    "Thro Yield", [Thro Yield £], "Availability", [Availabiilty £],
    "42 Activity", [42 Activity £], "36 Activity", [36 Activity £],
    "36 Speed", [36 Speed £], "42 Speed", [42 Speed £],
    "Bridge", [Total Vairance Calc], "Mill Accounts", [Total Variance Calc1], "Difference", [Difference]
)
```

## If the report pages don't open

The model was generated without Power BI Desktop to test it. If Desktop reports a problem with the report or a visual,
delete the contents of `DTA_Simplified.Report/definition/pages/` except `pages.json`, set `pages.json` to
`{"pageOrder": []}`, reopen, and build the pages yourself (about 10 minutes):

1. **Slicer:** `Dim_Week[Week No]`.
2. **Cards:** `Total Variance (Bridge) £`, `Total Variance (Mill Accounts) £`, `Unexplained Difference £`, `Reconciliation Status`.
3. **Waterfall chart:** Category = `Variance Bridge Steps[Step]`, Y-axis = `Bridge Step £`. Sort by Step, ascending.
4. **Table (page 2):** `Dim_Week[Week No]` plus the measures in the **Variance Bridge** folder.

Then send the error text to Claude so the generated report can be fixed.
