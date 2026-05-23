# Deployment Guide — Ticket Intelligence & Operations Monitoring Platform (TIOMP)

A complete one-click enterprise dashboard built in pure Excel VBA from `Apps.csv`.
This guide walks you from a fresh Excel install to a fully wired production dashboard
in under 5 minutes.

---

## 1. Prerequisites

| Requirement              | Minimum                                |
| ------------------------ | -------------------------------------- |
| Microsoft Excel          | 2016 / 2019 / 2021 / Microsoft 365     |
| Operating System         | Windows 10/11 (Outlook automation needs Windows) |
| Trust Center             | "Trust access to the VBA project object model" — **enabled** |
| Macro Settings           | "Disable macros with notification" or "Enable" |
| Apps.csv location        | Same folder as the workbook (or one level up) |

---

## 2. Folder Layout

```
ExcelDashbaord/
├── Apps.csv                  # Source dataset (already in the repo)
├── DEPLOYMENT.md             # This file
├── README.md                 # Architecture & feature overview
└── vba/                      # All VBA source modules to import
    ├── M01_Builder.bas       # Orchestrator + theme + logger
    ├── M02_DataEngine.bas    # CSV import, cleaning, derived columns, model
    ├── M03_KpiEngine.bas     # 36+ KPIs, AI alerts, recommendations
    ├── M04_DashboardUI.bas   # Premium dashboard shell, KPI cards
    ├── M05_Charts.bas        # Dark-themed chart engine
    ├── M06_Modules.bas       # 8 specialized analytics module sheets
    ├── M07_Interaction.bas   # Navigation, search, exports, login
    └── ThisWorkbook.cls      # Workbook event handlers
```

---

## 3. First-Time Setup (5 minutes)

### Step 1 — Create the workbook

1. Open Excel and create a new blank workbook.
2. Save it as **`TIOMP.xlsm`** (Excel Macro-Enabled Workbook) inside the
   same folder as `Apps.csv`.

### Step 2 — Open the VBA editor

Press **`Alt + F11`** to open the Visual Basic for Applications IDE.

### Step 3 — Import all VBA modules

In the VBA IDE menu choose **File → Import File…** and import every file from
the `vba/` folder, **in this exact order**:

1. `M01_Builder.bas`
2. `M02_DataEngine.bas`
3. `M03_KpiEngine.bas`
4. `M04_DashboardUI.bas`
5. `M05_Charts.bas`
6. `M06_Modules.bas`
7. `M07_Interaction.bas`

> **`ThisWorkbook.cls`** cannot be imported as a module. Instead, double-click
> the existing `ThisWorkbook` object in the VBA project tree on the left, then
> open `vba/ThisWorkbook.cls` in any text editor and copy/paste **only the code
> below the header attributes** (everything from `Option Explicit` onwards) into
> the existing `ThisWorkbook` code window.

### Step 4 — Save the workbook

Switch back to Excel and save (`Ctrl + S`). Confirm the format is `.xlsm`.

### Step 5 — Run the one-click build

Run the master macro using either method:

- **Method A:** Press `Alt + F8`, select **`BuildEnterpriseDashboard`** and click *Run*.
- **Method B:** From the VBA IDE place the cursor inside the `BuildEnterpriseDashboard`
  procedure in `M01_Builder` and press **`F5`**.

The build typically completes in **2–5 seconds** and produces:

- 16 sheets (Dashboard_Main + 7 analytics modules + 5 data + Settings + Logs + Hidden)
- 18 enterprise KPIs
- 8 hero KPI cards + 10 secondary slim KPIs on the main dashboard
- 13+ dark-themed charts (trend, donut, gauge, radar, scatter, heatmap, …)
- AI Alert Feed with rule-based smart alerts
- Live Activity Console with the latest 8 tickets
- Floating navigation rail on every analytics sheet
- Settings page with full admin controls

---

## 4. Day-to-Day Usage

### Refresh the dashboard

- Click the **REFRESH** button on the Dashboard_Main footer, or
- Run macro `RefreshAll`, or
- Reopen the workbook and accept the refresh prompt

The CSV is re-read, all KPIs and AI features are recomputed and every chart and
KPI card is repainted. The dashboard layout itself is preserved.

### Navigate between modules

Use the floating left-rail navigation. The active module is highlighted in
neon-cyan; clicking any other entry jumps directly to that page.

### Search any ticket

Click the **SEARCH** button (footer of Dashboard_Main) or run macro `SmartSearch`.
Enter any term — Issue ID, Title, Reporter, Location, Category, Severity or
Status. Results land on a freshly built `Search_Results` sheet with severity
and status pills.

### Export to PDF

Click **EXPORT PDF** on the dashboard footer. All 9 analytics sheets are
exported as a single landscape PDF in the workbook folder.

### Email the daily MIS

Click **EMAIL MIS**. If Outlook is installed, an HTML email is composed with
the executive KPI table prefilled. A `mailto:` fallback is used otherwise.

### Sign in / Switch role

From Settings → **SIGN IN**.
Default credentials (change them in `Hidden_Config` after first run):
- Username: `admin`
- Password: `admin`

The role pill in the header updates to reflect the current user/role.

### Backup & Archive

- **BACKUP** — saves a timestamped copy of the workbook
- **ARCHIVE CLOSED** — moves all closed tickets into a dated archive sheet

---

## 5. Customization

### Change SLA targets

Edit the constants near the top of `M01_Builder.bas`:

```vb
Public Const SLA_CRITICAL_H As Double = 4#
Public Const SLA_HIGH_H     As Double = 8#
Public Const SLA_MEDIUM_H   As Double = 24#
Public Const SLA_LOW_H      As Double = 48#
```

Then run `RefreshAll`.

### Change the theme palette

All colors live as `Public Const CLR_*` in `M01_Builder.bas`. Override
`CLR_BG`, `CLR_PANEL`, `CLR_ACCENT` etc. and rebuild.

### Change the source file

Edit `CSV_FILE_NAME` in `M01_Builder.bas`. The path resolver checks the
workbook folder, a `data/` subfolder, and the parent folder before falling
back to a file picker dialog.

### Override admin credentials

Open the `Hidden_Config` sheet (use macro `Sheets("Hidden_Config").Visible = True`
in the Immediate window once, edit values in `B3:B5`, then set Visible back
to `xlSheetVeryHidden`).

---

## 6. Troubleshooting

| Symptom                                | Cause / Fix                                            |
| -------------------------------------- | ------------------------------------------------------ |
| `Compile error: User-defined type not defined` | A module was imported out of order — reimport in the order shown in Step 3 |
| `Run-time error 5001: Apps.csv not found` | Place `Apps.csv` next to the workbook or in `data/` subfolder, or use the file picker prompt |
| All cards show `0`                     | Run `RefreshAll`. Check `Logs` sheet for the error    |
| Charts look light-themed               | The user toggled theme — click **THEME** to switch back |
| Outlook email fails                    | Outlook not installed — the macro falls back to `mailto:` automatically |
| Shapes overlap on smaller monitors     | Use `View → Zoom → Fit Selection`, or reduce the grid by editing `PaintCanvas` |
| `BuildEnterpriseDashboard` is slow     | Disable add-ins, close other workbooks, ensure CSV is on local disk (not network share) |

---

## 7. Security Notes

- The workbook stores credentials in plain text on `Hidden_Config`. For real
  production deployment, replace with hashed values via `M07_Interaction`.
- Set workbook structure protection: **Review → Protect Workbook** with a
  password to prevent users from unhiding the data sheets.
- Lock the VBA project: **VBA IDE → Tools → VBAProject Properties → Protection**
  and set a password.
- Optional: encrypt the file at rest via **File → Info → Protect Workbook →
  Encrypt with Password**.

---

## 8. Performance Optimization (Already Implemented)

- All CSV parsing happens in-memory via array reads (single shot writes)
- `ScreenUpdating`, `EnableEvents`, `Calculation` are toggled off during builds
- Aggregations use `Scripting.Dictionary` for O(1) lookups
- Chart styling reuses a single `StyleChartDark` helper across all chart objects
- The 8 analytics modules share KPI cards via a single `DrawCardRow` helper

---

## 9. Verifying the Build

After running `BuildEnterpriseDashboard`, verify these end-to-end:

1. Dashboard_Main shows 8 hero KPI cards + 10 slim KPIs (no `#REF!` or zeros if data is loaded)
2. Trend chart shows ticket volume + 3-day forecast (purple dashed line)
3. AI Alert Feed shows colored cards on the right (red Critical, amber High, blue Medium)
4. Live Activity Console shows 8 most recent tickets with colored severity dots
5. Side-rail navigation navigates between all 9 module sheets without errors
6. **SEARCH** button finds tickets by any term
7. **EXPORT PDF** produces a multi-sheet PDF
8. The `Logs` sheet shows entries for each phase (DATA, KPI, AI, UI, BUILD)

You're done. Welcome to the war room.
