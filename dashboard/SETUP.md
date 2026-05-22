# Ultra Premium Excel Dashboard - Setup Guide

A fully VBA-powered, corporate-style executive dashboard built on top of `a.csv` and `b.csv`.

---

## What you get

When you run the macro it builds 4 sheets in your workbook:

| Sheet | Purpose |
|-------|---------|
| **Dashboard** | Premium executive view (KPIs + 5 charts + header/footer) |
| **Data_A** | Raw `a.csv` import (formatted) |
| **Data_B** | Raw `b.csv` import (formatted, used for date trend) |
| **Calc** | Aggregations powering the visuals |

### Visual highlights
- Deep navy + gold corporate theme (`#0F1626` / `#D4AF37`)
- Brand bar with monogram, **LIVE** status pill, live date pill
- 6 KPI cards with accent stripes, icon circles, drop shadows
- Charts: Revenue by Category, Top 10 Products, Monthly Sales Trend (area + line overlay), Top 10 Cities, Rating Distribution donut with centered total
- Confidential footer with timestamp + bottom gold accent strip

---

## Setup (one-time, ~2 minutes)

### Step 1 - Get the files locally
Clone or download the repo so you have:
```
ExcelDashbaord/
  a.csv
  b.csv
  dashboard/
    DashboardBuilder.bas
    SETUP.md
```

### Step 2 - Create a macro-enabled workbook
1. Open Excel and create a **new blank workbook**.
2. Save it as **`Dashboard.xlsm`** in the **same folder as `a.csv` and `b.csv`** (i.e., the repo root). The `.xlsm` extension is required for macros.

### Step 3 - Import the VBA module
1. Press **`Alt + F11`** to open the VBA editor.
2. In the menu: **File -> Import File...**
3. Select **`dashboard/DashboardBuilder.bas`**.
4. You should now see a `DashboardBuilder` module in the Project Explorer (left panel).
5. Close the VBA editor.

### Step 4 - Run the dashboard build
1. Back in Excel, press **`Alt + F8`**.
2. Select **`BuildDashboard`**.
3. Click **Run**.

Wait ~5-15 seconds. You'll get a popup: `"Premium Dashboard built in X.X s."`

That's it. The Dashboard tab will be active and ready.

---

## If a.csv / b.csv aren't in the same folder

The macro first looks in the workbook's folder. If a file isn't found, a file picker opens automatically - just navigate to the CSV and click Open. It will do this once for `a.csv`, then once for `b.csv`.

---

## Re-running / refreshing

You can run `BuildDashboard` again anytime - it will fully rebuild all sheets, charts, and shapes from the latest CSV data.

To start completely fresh, run:
```
ResetWorkbook
```
This deletes all 4 generated sheets so you can rebuild cleanly.

---

## Customization tips

All the colors are inlined as `RGB(...)` calls so you can search-replace easily:

| Color use | Value |
|-----------|-------|
| Background navy | `RGB(15, 22, 38)` |
| Card / panel | `RGB(31, 42, 64)` |
| Border / line | `RGB(58, 72, 102)` |
| Gold accent | `RGB(212, 175, 55)` |
| Teal accent | `RGB(0, 201, 167)` |
| Coral accent | `RGB(255, 99, 132)` |
| Muted text | `RGB(157, 163, 180)` |

Other quick edits inside `DashboardBuilder.bas`:
- **Title text** - search for `EXECUTIVE COMMERCE DASHBOARD` (in `BuildHeader`)
- **Brand name** - search for `KIRO ANALYTICS`
- **Footer text** - in `BuildFooter`
- **Number formats** - search `NumberFormat = "$#,##0,K"` to change currency / scaling
- **Chart positions** - the ranges in `BuildCharts` (e.g. `B14:J29`) define each chart's grid area

---

## Troubleshooting

**"Macros are disabled"**
File -> Options -> Trust Center -> Trust Center Settings -> Macro Settings -> "Enable VBA macros" (or use a Trusted Location).

**"Cannot run the macro... may not be available in this workbook"**
You opened `.xlsx` (not macro-enabled). Save As `.xlsm`.

**File-not-found error during import**
The picker dialog will open - just point to `a.csv` and `b.csv` manually.

**Chart looks tiny / large**
Excel respects per-monitor DPI. Adjust the ranges in `BuildCharts` (e.g. widen `B14:J29` to `B14:K29`) and re-run.

**Date trend chart is empty**
This pulls from `b.csv` only (which has ISO `YYYY-MM-DD` dates). Make sure `b.csv` is loaded into `Data_B`.

---

## Files

- `DashboardBuilder.bas` - the entire VBA module (single file, ~900 lines)
- `SETUP.md` - this guide
