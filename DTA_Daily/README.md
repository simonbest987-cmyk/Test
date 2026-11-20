# DTA_Daily – Mill performance report from the daily tonnage export

A new Power BI project (PBIP) built on the **daily_tonnage_account.xlsx** export. Its four sheets are Tonnage Account,
Rolling Times, Delays and Speed Feedback. It uses the same calculations as the Daily Tonnage Account model, arranged as
a star schema. The existing Daily Tonnage Account project is not changed.

## Business rules applied

| Rule | How it is implemented |
|---|---|
| **Operational day = 07:00 to 07:00** | `Operational Date = Date(timestamp − 7h)`. A bar sawn at 06:57 on 19 Nov belongs to **18 Nov**, and one at 07:02 belongs to **19 Nov**. Rolling, delay and speed periods that cross 07:00 are **split**, and each part is counted on its own day. |
| **Financial weeks run Sunday to Saturday** | `Dim_Date[Week Commencing]` is the Sunday. Weeks are labelled `w/c 15 Nov 2026`. |
| Week / period / year **numbering** | **Not applied.** Neither the export nor the model shows the numbering rule. Add it to `Dim_Date` when available. |
| Product codes and mappings | Kept as exported (`Product Code`, `group descr`, `product`). `Product Key` only trims trailing spaces, so `"CR150      "` and `"CR150"` match across sheets. |
| Formulas | Same as the existing model's columns and measures. Each measure description names the model object it matches. |

## Open it

1. Create a folder, e.g. `C:\PBI\DTA_Exports\`, and put the daily export files in it. Any file whose name contains
   `daily_tonnage_account` and ends `.xlsx` is picked up, so give each day's file a unique name,
   e.g. `daily_tonnage_account 2026-11-20.xlsx`.
2. Unzip or clone this repo to a **short path** (e.g. `C:\PBI\Test\`) and open **`DTA_Daily\DTA_Daily.pbip`**.
3. If your folder is elsewhere: **Home → Transform data → Edit parameters → `ExportFolderPath`**.
   For a SharePoint folder, paste the full folder URL (`https://britishsteelltd.sharepoint.com/sites/NorthernMill-Finance/Shared Documents/<folder>/`).
   `SharePointSiteUrl` is already set to the same site as the existing model.
4. **Home → Refresh.**

Each export repeats the week to date. Overlapping files are de-duplicated:
- **Bars:** one row per Cast + Piece + saw time.
- **Delays, rolling times and speed periods:** one row per start time + product (+ area/location for delays). The newest file wins, so an open period gets its final end time.

## Model (star schema)

```
                 Dim_Date (Operational Date, Week Commencing, Week, Day, Calendar Month)
                    │ 1
   ┌───────────┬────┴──────┬──────────────┐
   * Fact_Bars  * Fact_Delays  * Fact_RollingTimes  * Fact_Speed
   └───────────┴────┬──────┴──────────────┘
                    │ 1
                 Dim_Product (Product Key, Product Group, Product Number)
```

* **8 relationships**, all many-to-one and single-direction. None are inactive, and there's no many-to-many.
* **Auto date/time is off.** `Dim_Date` is marked as the date table and covers whole Sunday–Saturday weeks.
* **The cost card stays on each bar** as it was exported. The cost on the day a bar was rolled is therefore kept, even if
  the cost card changes later.
* **No calculated columns or calculated tables.** Row logic is in Power Query; all aggregation is in measures.
  The raw numeric columns are hidden, and implicit measures are switched off.

## Measures (`_Measures`)

| Folder | Measures | Format |
|---|---|---|
| 1 Production | Bars Rolled, Bloom (t), Cut (t), Prime (t) | tonnes `#,0.0` |
| 2 Yield | 42 / 36 / Mill Through Yield %, each with **Std %** and **Var (pts)** | `0.0%`, pts `+0.0%` |
| 3 Cost | Bloom Consumption £, 42 Production Value £, Production Value £, 42/36 Scrap £, 42/36 Scale £, OSP £, Scrap & Scale £, **Bloom Cost (£/t)**, **Production Value (£/t)** | `£#,0`, `£#,0.00` |
| 4 Time & Delays | Rolling (h), Delay (h), Delay (hh:mm), Delay Events, Average Delay (min), Worked (h), Availability %, Delay % of Rolling, 42 / 36 Standard (h) | hours `#,0.0` |
| 5 Speed | Speed Logged (h), Actual Speed, **Target Speed**, Speed Var, Speed Var % | `#,0.0` |
| 6 Previous Week | Prime (t), Availability %, Delay (h), Production Value (£/t), Mill Through Yield %: **Prev Week**, **vs Prev Week**, **vs Prev Week %** or **(pts)** | as base |

**Actual vs target** is available where the export holds a target: the standard yields from the cost card, and `target_speed`.

**Not available from this export:**
- **Budget:** there is no annual plan in the export.
- **YTD:** needs the financial year rule.
- **The £ variance bridge:** needs the Operating_Costs cost-centre rates.
- **Finishing / defects yield:** the export has no defects column, so Prime = Total Output.

## Report pages

| Page | Contents |
|---|---|
| **Overview** | KPI cards: Prime t, Through Yield, Availability, £/t, Delay h. Prime trend by day, rolling vs delay hours by day, and a financial week comparison with the previous week. |
| **Yield & Cost** | 42/36 yield and variance cards, Bloom £/t, a product table (tonnes, actual vs std yield, £, £/t), cost drivers (scrap, scale, OSP), and a yield trend. |
| **Delays & Availability** | Cards; delay hours by area and by cause (largest first), by day and type, and the delay log with comments. |
| **Speed** | Actual vs target cards, actual vs target by product, hours by speed reason, and the speed log. |

Every page has **Financial week** and **Product** slicers.

## Check it: expected results for the sample file (w/c 15 Nov 2026)

With only the sample export in the folder, the report should show:

| Measure | Expected | | Measure | Expected |
|---|---:|---|---|---:|
| Bars Rolled | 105 | | Rolling (h) | 109.2 |
| Bloom (t) | 397.1 | | Delay (h) | 68.1 |
| Cut (t) | 391.1 | | Worked (h) | 41.1 |
| Prime (t) | 355.7 | | Availability % | 37.7% |
| 42 Mill Yield % / Std % | 98.5% / 98.5% | | Delay Events | 139 |
| 36 Mill Yield % / Std % | 91.0% / 92.8% | | Average Delay (min) | 29 |
| Mill Through Yield % / Std % | 89.6% / 91.4% | | Actual / Target Speed | 41.8 / 44.6 |
| Bloom Consumption £ | £177,607 | | 36 Scrap £ | £8,173 |
| Production Value £ | £253,280 | | 42 Scrap £ | £265 |
| Bloom Cost (£/t) | £447.28 | | OSP £ | £1,326 |
| Production Value (£/t) | £712.03 | | 42 / 36 Scale £ | £89 / £68 |

**By operational day:**

| Day | Prime (t) | Rolling (h) | Delay (h) | Availability % |
|---|---:|---:|---:|---:|
| Sun 15 Nov | – | 13.4 | 12.6 | 6.2% |
| Mon 16 Nov | – | 24.0 | 10.7 | 55.2% |
| Tue 17 Nov | – | 24.0 | 12.4 | 48.2% |
| Wed 18 Nov | 0.4 (1 bar, 06:57 Shift C) | 24.0 | 16.3 | 31.9% |
| Thu 19 Nov | 355.3 | 23.8 | 16.0 | 32.8% |

These figures were calculated independently from the sample file, applying the same rules.

**Things to note in the data:**
- **Availability is low** because many delays fall outside the rolling periods. For example, roll changes before the
  first rolling period started at 17:33 on 15 Nov. The formula matches the existing model (`Avail` = worked ÷ rolling hours).
  Please confirm that this is the intended definition for this data.
- **Delay areas include both `Finishing` and `Finishing Dept.`.** They are kept as exported, not merged.

## If something doesn't open

The project was generated and checked by script: DAX references, relationships, Power Query output columns, report fields
and JSON. It has **not** been opened in Power BI Desktop. If Desktop reports an error, send the error text.
