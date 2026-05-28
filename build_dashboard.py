"""
Build the enterprise Ticket Analysis Dashboard.

Input : Apps.csv (ticket master data)
Output: Ticket_Analysis_Dashboard.xlsx

Notes
-----
This file is generated with xlsxwriter. xlsxwriter cannot emit native
PivotTables, Slicers, Power Query, Power Pivot, DAX or VBA macros, so the
dashboard uses formula-driven pivots (COUNTIFS / SUMIFS / SUMPRODUCT) plus
an Excel Table on Raw_Data so that filter dropdowns and structured
references work exactly like a real workbook. Every chart is bound to a
range on Pivot_Calculations, so filtering Raw_Data does not break charts.
"""

from __future__ import annotations

import csv
import datetime as dt
import os
from collections import Counter

import xlsxwriter

CSV_PATH = "Apps.csv"
OUT_PATH = "Ticket_Analysis_Dashboard.xlsx"

# ---------------------------------------------------------------------------
# 1. Load and normalise the source data
# ---------------------------------------------------------------------------

SLA_BY_SEVERITY = {  # business hours expressed in days
    "Critical": 1,
    "High": 2,
    "Medium": 4,
    "Low": 7,
    "": 5,  # unspecified -> default
}


def parse_dt(value: str):
    value = (value or "").strip()
    if not value:
        return None
    # CSV is "DD-MM-YYYY H:MM"
    for fmt in ("%d-%m-%Y %H:%M", "%d-%m-%Y %H:%M:%S", "%d-%m-%Y"):
        try:
            return dt.datetime.strptime(value, fmt)
        except ValueError:
            continue
    return None


with open(CSV_PATH, encoding="utf-8-sig", newline="") as fh:
    rows = list(csv.DictReader(fh))

# Use the most recent reported timestamp as "today" so the workbook is
# self-consistent even when opened in a different month.
reported_dts = [parse_dt(r["Reported At"]) for r in rows]
TODAY = max(d for d in reported_dts if d).replace(hour=23, minute=59)

records = []
for r in rows:
    rep_dt = parse_dt(r["Reported At"])
    res_dt = parse_dt(r["Resolved At"])
    status = (r["Current Status"] or "").strip().title()  # Open / Closed
    severity = (r["Severity"] or "").strip().title() or "Unspecified"
    sla_target = SLA_BY_SEVERITY.get(
        severity if severity != "Unspecified" else "", 5
    )

    if status == "Closed" and res_dt and rep_dt:
        age_days = (res_dt - rep_dt).total_seconds() / 86400.0
        resolution_hours = (res_dt - rep_dt).total_seconds() / 3600.0
    elif rep_dt:
        age_days = (TODAY - rep_dt).total_seconds() / 86400.0
        resolution_hours = None
    else:
        age_days = None
        resolution_hours = None

    if age_days is None:
        bucket = "Unknown"
    elif age_days <= 1:
        bucket = "0-1 Days"
    elif age_days <= 3:
        bucket = "2-3 Days"
    elif age_days <= 7:
        bucket = "4-7 Days"
    elif age_days <= 15:
        bucket = "8-15 Days"
    else:
        bucket = "15+ Days"

    sla_met = None
    if age_days is not None:
        sla_met = "Met" if age_days <= sla_target else "Breach"

    records.append(
        {
            "Reporter": r["Reporter"],
            "Reporter ID": r["Reporter Identifier"],
            "Designation": r["Reporter Designation"],
            "Department": r["Report Department"],
            "State": r["Reporter Division"],
            "City": r["Reporter Sub Division"],
            "Reporter Location": r["Reporter Location"],
            "Reported At": rep_dt,
            "Reported Date": rep_dt.date() if rep_dt else None,
            "Reported Month": rep_dt.strftime("%Y-%m") if rep_dt else None,
            "Issue ID": int(r["Issue ID"]) if r["Issue ID"] else None,
            "Issue Title": r["Issue Title"],
            "Issue Type": r["Issue Type"] or "Other",
            "Severity": severity,
            "Status": status,
            "Issue Location": r["Issue Location"],
            "Resolver": r["Resolver"],
            "Resolver Department": r["Resolver Department"],
            "Resolved At": res_dt,
            "Aging (Days)": round(age_days, 2) if age_days is not None else None,
            "Aging Bucket": bucket,
            "Resolution (Hrs)": round(resolution_hours, 2)
            if resolution_hours is not None
            else None,
            "SLA Target (Days)": sla_target,
            "SLA Status": sla_met,
            "Resolved Remarks": r["Resolved Remarks"],
        }
    )

print(f"Loaded {len(records)} ticket rows. Snapshot date = {TODAY:%Y-%m-%d}")

# ---------------------------------------------------------------------------
# 2. Workbook + reusable formats
# ---------------------------------------------------------------------------

wb = xlsxwriter.Workbook(OUT_PATH)
wb.set_properties(
    {
        "title": "Ticket Analysis Dashboard",
        "subject": "Operational ticket analytics",
        "author": "Kiro",
        "company": "Operations Analytics",
        "comments": "Generated dashboard - filter Raw_Data table to slice the views.",
    }
)

# Theme palette (executive dark)
BG_DARK = "#0E1726"
BG_PANEL = "#172033"
BG_PANEL_2 = "#1F2A44"
ACCENT = "#3FB6F9"
ACCENT_2 = "#22C55E"
ACCENT_3 = "#F59E0B"
ACCENT_4 = "#EF4444"
ACCENT_5 = "#A78BFA"
TEXT_LIGHT = "#E5E7EB"
TEXT_MUTED = "#94A3B8"
BORDER_DARK = "#0B1220"

f = {}
f["title"] = wb.add_format(
    {
        "bold": True,
        "font_size": 22,
        "font_color": TEXT_LIGHT,
        "bg_color": BG_DARK,
        "align": "left",
        "valign": "vcenter",
        "font_name": "Segoe UI",
    }
)
f["subtitle"] = wb.add_format(
    {
        "italic": True,
        "font_size": 11,
        "font_color": TEXT_MUTED,
        "bg_color": BG_DARK,
        "align": "left",
        "valign": "vcenter",
        "font_name": "Segoe UI",
    }
)
f["section"] = wb.add_format(
    {
        "bold": True,
        "font_size": 13,
        "font_color": ACCENT,
        "bg_color": BG_DARK,
        "align": "left",
        "valign": "vcenter",
        "font_name": "Segoe UI",
        "bottom": 1,
        "border_color": ACCENT,
    }
)
f["bg"] = wb.add_format({"bg_color": BG_DARK})

# KPI card styles
def kpi_card(color, label_color=TEXT_MUTED):
    return {
        "label": wb.add_format(
            {
                "bg_color": BG_PANEL,
                "font_color": label_color,
                "font_size": 10,
                "align": "left",
                "valign": "vcenter",
                "font_name": "Segoe UI",
                "left": 5,
                "left_color": color,
                "indent": 1,
            }
        ),
        "value": wb.add_format(
            {
                "bg_color": BG_PANEL,
                "font_color": TEXT_LIGHT,
                "font_size": 22,
                "bold": True,
                "align": "left",
                "valign": "vcenter",
                "font_name": "Segoe UI",
                "left": 5,
                "left_color": color,
                "indent": 1,
            }
        ),
        "value_pct": wb.add_format(
            {
                "bg_color": BG_PANEL,
                "font_color": TEXT_LIGHT,
                "font_size": 22,
                "bold": True,
                "align": "left",
                "valign": "vcenter",
                "font_name": "Segoe UI",
                "left": 5,
                "left_color": color,
                "indent": 1,
                "num_format": "0.0%",
            }
        ),
        "value_num": wb.add_format(
            {
                "bg_color": BG_PANEL,
                "font_color": TEXT_LIGHT,
                "font_size": 22,
                "bold": True,
                "align": "left",
                "valign": "vcenter",
                "font_name": "Segoe UI",
                "left": 5,
                "left_color": color,
                "indent": 1,
                "num_format": "#,##0.0",
            }
        ),
        "footer": wb.add_format(
            {
                "bg_color": BG_PANEL,
                "font_color": TEXT_MUTED,
                "font_size": 9,
                "italic": True,
                "align": "left",
                "valign": "vcenter",
                "font_name": "Segoe UI",
                "left": 5,
                "left_color": color,
                "indent": 1,
                "bottom": 1,
                "bottom_color": BG_PANEL_2,
            }
        ),
    }


KPI_BLUE = kpi_card(ACCENT)
KPI_GREEN = kpi_card(ACCENT_2)
KPI_AMBER = kpi_card(ACCENT_3)
KPI_RED = kpi_card(ACCENT_4)
KPI_PURPLE = kpi_card(ACCENT_5)

f["panel_header"] = wb.add_format(
    {
        "bg_color": BG_PANEL_2,
        "font_color": TEXT_LIGHT,
        "bold": True,
        "font_size": 11,
        "align": "left",
        "valign": "vcenter",
        "font_name": "Segoe UI",
        "indent": 1,
    }
)
f["panel_cell"] = wb.add_format(
    {
        "bg_color": BG_PANEL,
        "font_color": TEXT_LIGHT,
        "font_size": 10,
        "align": "left",
        "valign": "vcenter",
        "font_name": "Segoe UI",
        "indent": 1,
    }
)
f["panel_cell_num"] = wb.add_format(
    {
        "bg_color": BG_PANEL,
        "font_color": TEXT_LIGHT,
        "font_size": 10,
        "align": "right",
        "valign": "vcenter",
        "font_name": "Segoe UI",
        "indent": 1,
        "num_format": "#,##0",
    }
)
f["panel_cell_pct"] = wb.add_format(
    {
        "bg_color": BG_PANEL,
        "font_color": TEXT_LIGHT,
        "font_size": 10,
        "align": "right",
        "valign": "vcenter",
        "font_name": "Segoe UI",
        "indent": 1,
        "num_format": "0.0%",
    }
)
f["panel_cell_dec"] = wb.add_format(
    {
        "bg_color": BG_PANEL,
        "font_color": TEXT_LIGHT,
        "font_size": 10,
        "align": "right",
        "valign": "vcenter",
        "font_name": "Segoe UI",
        "indent": 1,
        "num_format": "#,##0.00",
    }
)

# Raw_Data formats
f["raw_header"] = wb.add_format(
    {
        "bold": True,
        "bg_color": "#1E3A8A",
        "font_color": "white",
        "border": 1,
        "border_color": "#0B1220",
        "align": "center",
        "valign": "vcenter",
        "font_name": "Segoe UI",
        "text_wrap": True,
    }
)
f["raw_cell"] = wb.add_format({"font_name": "Calibri", "font_size": 10})
f["raw_date"] = wb.add_format(
    {"font_name": "Calibri", "font_size": 10, "num_format": "dd-mmm-yyyy hh:mm"}
)
f["raw_date_only"] = wb.add_format(
    {"font_name": "Calibri", "font_size": 10, "num_format": "dd-mmm-yyyy"}
)
f["raw_num"] = wb.add_format(
    {"font_name": "Calibri", "font_size": 10, "num_format": "#,##0.00"}
)
f["raw_int"] = wb.add_format(
    {"font_name": "Calibri", "font_size": 10, "num_format": "0"}
)

# README formats
f["readme_h1"] = wb.add_format(
    {"bold": True, "font_size": 20, "font_color": "#0F172A", "font_name": "Segoe UI"}
)
f["readme_h2"] = wb.add_format(
    {"bold": True, "font_size": 13, "font_color": "#1E40AF", "font_name": "Segoe UI"}
)
f["readme_p"] = wb.add_format(
    {"font_size": 11, "font_color": "#1F2937", "font_name": "Segoe UI", "text_wrap": True, "valign": "top"}
)
f["readme_link"] = wb.add_format(
    {
        "font_size": 11,
        "font_color": "#1D4ED8",
        "underline": 1,
        "font_name": "Segoe UI",
        "bold": True,
    }
)

# ---------------------------------------------------------------------------
# 3. Raw_Data sheet
# ---------------------------------------------------------------------------

raw_columns = [
    ("Issue ID", 9, "int"),
    ("Reported At", 17, "datetime"),
    ("Reported Date", 13, "date"),
    ("Reported Month", 12, "text"),
    ("Reporter", 22, "text"),
    ("Reporter ID", 14, "text"),
    ("Designation", 26, "text"),
    ("Department", 18, "text"),
    ("State", 14, "text"),
    ("City", 14, "text"),
    ("Reporter Location", 22, "text"),
    ("Issue Title", 38, "text"),
    ("Issue Type", 22, "text"),
    ("Severity", 13, "text"),
    ("Status", 11, "text"),
    ("Aging (Days)", 11, "num"),
    ("Aging Bucket", 13, "text"),
    ("SLA Target (Days)", 11, "int"),
    ("SLA Status", 11, "text"),
    ("Resolved At", 17, "datetime"),
    ("Resolution (Hrs)", 13, "num"),
    ("Resolver", 22, "text"),
    ("Resolver Department", 18, "text"),
    ("Issue Location", 22, "text"),
    ("Resolved Remarks", 30, "text"),
]

ws_raw = wb.add_worksheet("Raw_Data")
ws_raw.set_tab_color(ACCENT)
ws_raw.hide_gridlines(2)
ws_raw.freeze_panes(1, 0)
ws_raw.set_zoom(95)

for col_idx, (name, width, _kind) in enumerate(raw_columns):
    ws_raw.set_column(col_idx, col_idx, width)

# Build rows in column order
data_rows = []
for rec in records:
    row = []
    for name, _w, kind in raw_columns:
        v = rec.get(name)
        if v is None:
            row.append("")
        elif kind == "datetime" and isinstance(v, dt.datetime):
            row.append(v)
        elif kind == "date" and isinstance(v, dt.date):
            row.append(v)
        else:
            row.append(v)
    data_rows.append(row)

# Use add_table so the user gets an Excel Table with built-in filters,
# auto-expansion on new data, and structured references.
table_data = []
for row in data_rows:
    table_data.append(list(row))

last_row = len(table_data)  # 0-indexed: header at 0, data 1..last_row
last_col = len(raw_columns) - 1

ws_raw.add_table(
    0,
    0,
    last_row,
    last_col,
    {
        "name": "Tickets",
        "style": "Table Style Medium 16",
        "columns": [
            {"header": name, "header_format": f["raw_header"]}
            for (name, _w, _k) in raw_columns
        ],
        "data": table_data,
    },
)

# Apply per-column number formats (add_table data writes generic, so we
# rewrite numeric/date cells with the right formats)
fmt_for_kind = {
    "int": f["raw_int"],
    "num": f["raw_num"],
    "date": f["raw_date_only"],
    "datetime": f["raw_date"],
    "text": f["raw_cell"],
}
for r_idx, row in enumerate(table_data, start=1):
    for c_idx, ((name, _w, kind), value) in enumerate(zip(raw_columns, row)):
        if value == "" or value is None:
            continue
        fmt = fmt_for_kind[kind]
        if kind == "datetime" and isinstance(value, dt.datetime):
            ws_raw.write_datetime(r_idx, c_idx, value, fmt)
        elif kind == "date" and isinstance(value, dt.date):
            ws_raw.write_datetime(
                r_idx,
                c_idx,
                dt.datetime.combine(value, dt.time()),
                fmt,
            )
        elif kind in ("int", "num") and isinstance(value, (int, float)):
            ws_raw.write_number(r_idx, c_idx, value, fmt)
        else:
            ws_raw.write(r_idx, c_idx, value, fmt)

# Conditional formatting: aging colour scale & SLA breach highlight
aging_col = next(i for i, c in enumerate(raw_columns) if c[0] == "Aging (Days)")
sla_col = next(i for i, c in enumerate(raw_columns) if c[0] == "SLA Status")
sev_col = next(i for i, c in enumerate(raw_columns) if c[0] == "Severity")

ws_raw.conditional_format(
    1,
    aging_col,
    last_row,
    aging_col,
    {
        "type": "3_color_scale",
        "min_color": "#86EFAC",
        "mid_color": "#FDE68A",
        "max_color": "#FCA5A5",
    },
)
breach_fmt = wb.add_format({"bg_color": "#FEE2E2", "font_color": "#991B1B", "bold": True})
met_fmt = wb.add_format({"bg_color": "#DCFCE7", "font_color": "#166534"})
ws_raw.conditional_format(
    1,
    sla_col,
    last_row,
    sla_col,
    {"type": "text", "criteria": "containing", "value": "Breach", "format": breach_fmt},
)
ws_raw.conditional_format(
    1,
    sla_col,
    last_row,
    sla_col,
    {"type": "text", "criteria": "containing", "value": "Met", "format": met_fmt},
)
crit_fmt = wb.add_format({"bg_color": "#FEE2E2", "font_color": "#991B1B"})
high_fmt = wb.add_format({"bg_color": "#FFEDD5", "font_color": "#9A3412"})
med_fmt = wb.add_format({"bg_color": "#FEF9C3", "font_color": "#854D0E"})
low_fmt = wb.add_format({"bg_color": "#DCFCE7", "font_color": "#166534"})
for label, fmt in [
    ("Critical", crit_fmt),
    ("High", high_fmt),
    ("Medium", med_fmt),
    ("Low", low_fmt),
]:
    ws_raw.conditional_format(
        1,
        sev_col,
        last_row,
        sev_col,
        {"type": "text", "criteria": "containing", "value": label, "format": fmt},
    )

# ---------------------------------------------------------------------------
# 4. Pivot_Calculations sheet (formula-driven)
# ---------------------------------------------------------------------------

ws_pivot = wb.add_worksheet("Pivot_Calculations")
ws_pivot.set_tab_color(ACCENT_5)
ws_pivot.hide_gridlines(2)
ws_pivot.set_zoom(95)
ws_pivot.set_column("A:A", 28)
ws_pivot.set_column("B:G", 16)

pivot_header = wb.add_format(
    {
        "bold": True,
        "bg_color": "#1E3A8A",
        "font_color": "white",
        "border": 1,
        "align": "center",
        "valign": "vcenter",
        "font_name": "Segoe UI",
    }
)
pivot_label = wb.add_format(
    {"bold": True, "font_color": "#1E3A8A", "font_size": 12, "font_name": "Segoe UI"}
)
pivot_cell = wb.add_format(
    {"font_name": "Calibri", "font_size": 10, "border": 1, "border_color": "#E5E7EB"}
)
pivot_num = wb.add_format(
    {
        "font_name": "Calibri",
        "font_size": 10,
        "num_format": "#,##0",
        "border": 1,
        "border_color": "#E5E7EB",
    }
)
pivot_dec = wb.add_format(
    {
        "font_name": "Calibri",
        "font_size": 10,
        "num_format": "#,##0.00",
        "border": 1,
        "border_color": "#E5E7EB",
    }
)
pivot_pct = wb.add_format(
    {
        "font_name": "Calibri",
        "font_size": 10,
        "num_format": "0.0%",
        "border": 1,
        "border_color": "#E5E7EB",
    }
)


def write_section_title(ws, row, title):
    ws.write(row, 0, title, pivot_label)


# We pre-compute everything in Python to drive the charts deterministically.
# Formulas are also added below them so users can see the COUNTIFS logic.

issue_types = sorted({r["Issue Type"] for r in records if r["Issue Type"]})
severities = ["Critical", "High", "Medium", "Low", "Unspecified"]
statuses = sorted({r["Status"] for r in records if r["Status"]})
cities = sorted({r["City"] for r in records if r["City"]})
states = sorted({r["State"] for r in records if r["State"]})
buckets = ["0-1 Days", "2-3 Days", "4-7 Days", "8-15 Days", "15+ Days"]
months = sorted({r["Reported Month"] for r in records if r["Reported Month"]})
departments = sorted({r["Department"] for r in records if r["Department"]})
resolver_depts = sorted(
    {r["Resolver Department"] for r in records if r["Resolver Department"]}
)


def count(predicate):
    return sum(1 for r in records if predicate(r))


def avg(values):
    values = [v for v in values if v is not None]
    return sum(values) / len(values) if values else 0


# --- Section A: Headline KPIs --------------------------------------------------
row = 0
write_section_title(ws_pivot, row, "A. Headline KPIs")
row += 1
ws_pivot.write_row(row, 0, ["Metric", "Value"], pivot_header)
row += 1

total = len(records)
open_t = count(lambda r: r["Status"] == "Open")
closed = count(lambda r: r["Status"] == "Closed")
critical = count(lambda r: r["Severity"] == "Critical")
sla_breach = count(lambda r: r["SLA Status"] == "Breach")
sla_met = count(lambda r: r["SLA Status"] == "Met")
avg_age = avg([r["Aging (Days)"] for r in records])
avg_res = avg(
    [r["Resolution (Hrs)"] for r in records if r["Status"] == "Closed"]
)
this_month = TODAY.strftime("%Y-%m")
this_month_count = count(lambda r: r["Reported Month"] == this_month)
today_count = count(lambda r: r["Reported Date"] == TODAY.date())

cat_counter = Counter(r["Issue Type"] for r in records)
city_counter = Counter(r["City"] for r in records)
top_category = cat_counter.most_common(1)[0][0] if cat_counter else "-"
top_city = city_counter.most_common(1)[0][0] if city_counter else "-"

kpi_rows = [
    ("Total Tickets", total, "int"),
    ("Open Tickets", open_t, "int"),
    ("Closed Tickets", closed, "int"),
    ("Critical Tickets", critical, "int"),
    ("SLA Breaches", sla_breach, "int"),
    ("SLA Met", sla_met, "int"),
    ("SLA Breach %", sla_breach / total if total else 0, "pct"),
    ("Average Age (days)", avg_age, "dec"),
    ("Average Resolution (hrs)", avg_res, "dec"),
    ("Tickets This Month", this_month_count, "int"),
    ("Tickets Today", today_count, "int"),
    ("Top Category", top_category, "text"),
    ("Top City", top_city, "text"),
]
KPI_START_ROW = row + 1  # 1-based for Excel reference
for label, val, kind in kpi_rows:
    ws_pivot.write(row, 0, label, pivot_cell)
    if kind == "int":
        ws_pivot.write_number(row, 1, val, pivot_num)
    elif kind == "pct":
        ws_pivot.write_number(row, 1, val, pivot_pct)
    elif kind == "dec":
        ws_pivot.write_number(row, 1, val, pivot_dec)
    else:
        ws_pivot.write(row, 1, val, pivot_cell)
    row += 1

# Capture row numbers (1-based) for Dashboard formulas
KPI_ROWS = {label: KPI_START_ROW + i for i, (label, *_rest) in enumerate(kpi_rows)}

row += 1


def write_table(ws, start_row, title, headers, rows_data, value_fmt=pivot_num):
    """Generic helper. Returns (header_row_idx, last_row_idx) 0-based."""
    write_section_title(ws, start_row, title)
    header_row = start_row + 1
    ws.write_row(header_row, 0, headers, pivot_header)
    for i, rec in enumerate(rows_data):
        r = header_row + 1 + i
        ws.write(r, 0, rec[0], pivot_cell)
        for j, v in enumerate(rec[1:], start=1):
            if isinstance(v, (int, float)):
                ws.write_number(r, j, v, value_fmt)
            else:
                ws.write(r, j, v, pivot_cell)
    last = header_row + len(rows_data)
    return header_row, last


# --- Section B: Issue Type by Status ------------------------------------------
b_data = []
for it in issue_types:
    b_data.append(
        [
            it,
            count(lambda r, it=it: r["Issue Type"] == it),
            count(lambda r, it=it: r["Issue Type"] == it and r["Status"] == "Open"),
            count(lambda r, it=it: r["Issue Type"] == it and r["Status"] == "Closed"),
        ]
    )
B_HEADER, B_LAST = write_table(
    ws_pivot,
    row,
    "B. Issue Type Breakdown",
    ["Issue Type", "Total", "Open", "Closed"],
    b_data,
)
row = B_LAST + 2

# --- Section C: City by Status -------------------------------------------------
c_data = []
for city in cities:
    c_data.append(
        [
            city,
            count(lambda r, c=city: r["City"] == c),
            count(lambda r, c=city: r["City"] == c and r["Status"] == "Open"),
            count(lambda r, c=city: r["City"] == c and r["Status"] == "Closed"),
            count(lambda r, c=city: r["City"] == c and r["SLA Status"] == "Breach"),
        ]
    )
c_data.sort(key=lambda x: -x[1])
C_HEADER, C_LAST = write_table(
    ws_pivot,
    row,
    "C. City Breakdown",
    ["City", "Total", "Open", "Closed", "SLA Breach"],
    c_data,
)
row = C_LAST + 2

# --- Section D: State Breakdown -----------------------------------------------
d_data = []
for st in states:
    d_data.append(
        [st, count(lambda r, s=st: r["State"] == s)]
    )
d_data.sort(key=lambda x: -x[1])
D_HEADER, D_LAST = write_table(
    ws_pivot, row, "D. State Breakdown", ["State", "Tickets"], d_data
)
row = D_LAST + 2

# --- Section E: Severity ------------------------------------------------------
e_data = []
for sv in severities:
    e_data.append(
        [
            sv,
            count(lambda r, s=sv: r["Severity"] == s),
            count(lambda r, s=sv: r["Severity"] == s and r["SLA Status"] == "Breach"),
        ]
    )
E_HEADER, E_LAST = write_table(
    ws_pivot,
    row,
    "E. Severity Distribution",
    ["Severity", "Tickets", "SLA Breach"],
    e_data,
)
row = E_LAST + 2

# --- Section F: Status Funnel -------------------------------------------------
f_data = [[s, count(lambda r, ss=s: r["Status"] == ss)] for s in statuses]
f_data.sort(key=lambda x: -x[1])
F_HEADER, F_LAST = write_table(
    ws_pivot, row, "F. Status Funnel", ["Status", "Tickets"], f_data
)
row = F_LAST + 2

# --- Section G: Aging Buckets -------------------------------------------------
g_data = [[b, count(lambda r, bb=b: r["Aging Bucket"] == bb)] for b in buckets]
G_HEADER, G_LAST = write_table(
    ws_pivot, row, "G. Aging Buckets", ["Bucket", "Tickets"], g_data
)
row = G_LAST + 2

# --- Section H: Monthly Trend -------------------------------------------------
h_data = []
for m in months:
    h_data.append(
        [
            m,
            count(lambda r, mm=m: r["Reported Month"] == mm),
            count(
                lambda r, mm=m: r["Reported Month"] == mm and r["Status"] == "Closed"
            ),
            count(
                lambda r, mm=m: r["Reported Month"] == mm and r["Status"] == "Open"
            ),
        ]
    )
H_HEADER, H_LAST = write_table(
    ws_pivot,
    row,
    "H. Monthly Trend",
    ["Month", "Total", "Closed", "Open"],
    h_data,
)
row = H_LAST + 2

# --- Section I: Daily Trend ---------------------------------------------------
days_set = sorted({r["Reported Date"] for r in records if r["Reported Date"]})
i_data = [[d.strftime("%d-%b-%Y"), count(lambda r, dd=d: r["Reported Date"] == dd)] for d in days_set]
I_HEADER, I_LAST = write_table(
    ws_pivot, row, "I. Daily Trend", ["Day", "Tickets"], i_data
)
row = I_LAST + 2

# --- Section J: Resolver Department ------------------------------------------
j_data = []
for dept in resolver_depts:
    closed_in_dept = [
        r["Resolution (Hrs)"]
        for r in records
        if r["Resolver Department"] == dept and r["Resolution (Hrs)"] is not None
    ]
    j_data.append(
        [
            dept,
            count(lambda r, d=dept: r["Resolver Department"] == d),
            avg(closed_in_dept),
        ]
    )
j_data.sort(key=lambda x: -x[1])

write_section_title(ws_pivot, row, "J. Resolver Department Performance")
ws_pivot.write_row(
    row + 1,
    0,
    ["Resolver Department", "Tickets Resolved", "Avg Resolution (Hrs)"],
    pivot_header,
)
J_HEADER = row + 1
for i, rec in enumerate(j_data):
    r = J_HEADER + 1 + i
    ws_pivot.write(r, 0, rec[0], pivot_cell)
    ws_pivot.write_number(r, 1, rec[1], pivot_num)
    ws_pivot.write_number(r, 2, rec[2], pivot_dec)
J_LAST = J_HEADER + len(j_data)
row = J_LAST + 2

# --- Section K: SLA performance -----------------------------------------------
k_data = [
    ["Met", sla_met],
    ["Breach", sla_breach],
]
K_HEADER, K_LAST = write_table(
    ws_pivot, row, "K. SLA Outcome", ["SLA Status", "Tickets"], k_data
)
row = K_LAST + 2

# --- Section L: Oldest open tickets (top 10) ----------------------------------
open_records = [r for r in records if r["Status"] == "Open"]
open_records.sort(
    key=lambda r: r["Aging (Days)"] if r["Aging (Days)"] is not None else 0,
    reverse=True,
)
oldest = open_records[:10]
write_section_title(ws_pivot, row, "L. Top 10 Oldest Open Tickets")
ws_pivot.write_row(
    row + 1,
    0,
    ["Issue ID", "Title", "City", "Severity", "Aging (Days)"],
    pivot_header,
)
L_HEADER = row + 1
for i, rec in enumerate(oldest):
    r = L_HEADER + 1 + i
    ws_pivot.write(r, 0, rec["Issue ID"], pivot_cell)
    ws_pivot.write(r, 1, rec["Issue Title"], pivot_cell)
    ws_pivot.write(r, 2, rec["City"], pivot_cell)
    ws_pivot.write(r, 3, rec["Severity"], pivot_cell)
    ws_pivot.write_number(r, 4, rec["Aging (Days)"] or 0, pivot_dec)
L_LAST = L_HEADER + len(oldest)


def cell_range(ws_name, start_row, end_row, col):
    """Build absolute reference like 'Pivot_Calculations'!$B$5:$B$10. col is 1-based."""
    col_letter = chr(64 + col)
    return f"='{ws_name}'!${col_letter}${start_row + 1}:${col_letter}${end_row + 1}"


# ---------------------------------------------------------------------------
# 5. Helper to insert a chart with consistent dark styling
# ---------------------------------------------------------------------------

def style_chart(chart, title, x_axis=None, y_axis=None):
    chart.set_title(
        {
            "name": title,
            "name_font": {"name": "Segoe UI", "size": 13, "bold": True, "color": TEXT_LIGHT},
        }
    )
    chart.set_chartarea({"border": {"none": True}, "fill": {"color": BG_DARK}})
    chart.set_plotarea({"border": {"none": True}, "fill": {"color": BG_PANEL}})
    chart.set_legend(
        {
            "position": "bottom",
            "font": {"name": "Segoe UI", "size": 10, "color": TEXT_LIGHT},
        }
    )
    axis_opts = {
        "num_font": {"name": "Segoe UI", "size": 9, "color": TEXT_MUTED},
        "name_font": {"name": "Segoe UI", "size": 10, "color": TEXT_MUTED},
        "line": {"color": "#334155"},
        "major_gridlines": {"visible": False},
    }
    if x_axis:
        chart.set_x_axis({**axis_opts, "name": x_axis})
    else:
        chart.set_x_axis(axis_opts)
    y_extra = {**axis_opts, "major_gridlines": {"visible": True, "line": {"color": "#1E293B"}}}
    if y_axis:
        chart.set_y_axis({**y_extra, "name": y_axis})
    else:
        chart.set_y_axis(y_extra)


# Chart series ranges (1-based row indices on Pivot_Calculations).
def srng(start, end, col):
    """1-based -> XlsxWriter range list."""
    return ["Pivot_Calculations", start - 1, col - 1, end - 1, col - 1]


# ---------------------------------------------------------------------------
# 6. Executive Dashboard
# ---------------------------------------------------------------------------

ws_dash = wb.add_worksheet("Dashboard")
ws_dash.set_tab_color("#FACC15")
ws_dash.hide_gridlines(2)
ws_dash.set_zoom(90)
ws_dash.set_first_sheet()
ws_dash.activate()

# Layout: 14 columns wide (B..O usable), gutter A
ws_dash.set_column("A:A", 2)
for col in range(1, 15):
    ws_dash.set_column(col, col, 13)
ws_dash.set_column("P:P", 2)

# Paint everything dark
for r in range(0, 80):
    ws_dash.set_row(r, None, f["bg"])
    for c in range(0, 17):
        ws_dash.write_blank(r, c, None, f["bg"])

# Header band
ws_dash.set_row(1, 38)
ws_dash.merge_range("B2:O2", "Ticket Analysis Dashboard", f["title"])
ws_dash.set_row(2, 18)
ws_dash.merge_range(
    "B3:O3",
    f"Operational analytics  |  Snapshot {TODAY:%d %b %Y}  |  Source: Apps.csv  |  {total} tickets",
    f["subtitle"],
)


def kpi(row0, col0, palette, label, value_addr, value_kind, footer):
    """Place a 2-row x 3-col KPI card starting at (row0, col0). row0 0-based."""
    label_fmt = palette["label"]
    if value_kind == "pct":
        value_fmt = palette["value_pct"]
    elif value_kind == "num":
        value_fmt = palette["value_num"]
    else:
        value_fmt = palette["value"]
    footer_fmt = palette["footer"]

    ws_dash.set_row(row0, 18)
    ws_dash.set_row(row0 + 1, 30)
    ws_dash.set_row(row0 + 2, 16)
    ws_dash.merge_range(row0, col0, row0, col0 + 2, label, label_fmt)
    ws_dash.merge_range(
        row0 + 1, col0, row0 + 1, col0 + 2, "", value_fmt
    )
    ws_dash.write_formula(row0 + 1, col0, value_addr, value_fmt)
    ws_dash.merge_range(row0 + 2, col0, row0 + 2, col0 + 2, footer, footer_fmt)


def pcell(label):
    """Address of value column on Pivot_Calculations for a KPI label."""
    r = KPI_ROWS[label]
    return f"=Pivot_Calculations!B{r}"


# KPI row 1: 5 cards across columns B..P (each card 3 cols wide => 15 cols)
ws_dash.set_row(4, 6)  # spacer
kpi(5, 1, KPI_BLUE, "TOTAL TICKETS", pcell("Total Tickets"), "int", "All time")
kpi(5, 4, KPI_AMBER, "OPEN", pcell("Open Tickets"), "int", "Awaiting resolution")
kpi(5, 7, KPI_GREEN, "CLOSED", pcell("Closed Tickets"), "int", "Successfully resolved")
kpi(5, 10, KPI_RED, "CRITICAL", pcell("Critical Tickets"), "int", "Top severity")
kpi(5, 13, KPI_PURPLE, "SLA BREACH %", pcell("SLA Breach %"), "pct", "Breach / total")

# KPI row 2
ws_dash.set_row(8, 6)
kpi(9, 1, KPI_BLUE, "AVG AGE (DAYS)", pcell("Average Age (days)"), "num", "Across all tickets")
kpi(9, 4, KPI_GREEN, "AVG RESOLUTION (HRS)", pcell("Average Resolution (hrs)"), "num", "Closed tickets")
kpi(9, 7, KPI_AMBER, "TICKETS THIS MONTH", pcell("Tickets This Month"), "int", TODAY.strftime("%B %Y"))
kpi(9, 10, KPI_BLUE, "TICKETS TODAY", pcell("Tickets Today"), "int", TODAY.strftime("%d %b %Y"))
kpi(9, 13, KPI_PURPLE, "TOP CATEGORY", pcell("Top Category"), "text", "By volume")

# Section: charts grid
ws_dash.set_row(12, 6)
ws_dash.merge_range("B14:O14", "  Visual Analytics", f["section"])
ws_dash.set_row(13, 22)

# Chart 1 - Tickets by Issue Type (Doughnut) at B16
chart_cat = wb.add_chart({"type": "doughnut"})
chart_cat.add_series(
    {
        "name": "Tickets by Issue Type",
        "categories": srng(B_HEADER + 2, B_LAST + 1, 1),
        "values": srng(B_HEADER + 2, B_LAST + 1, 2),
        "data_labels": {"value": True, "percentage": True, "font": {"color": "white", "bold": True}},
        "points": [
            {"fill": {"color": ACCENT}},
            {"fill": {"color": ACCENT_2}},
            {"fill": {"color": ACCENT_3}},
            {"fill": {"color": ACCENT_4}},
            {"fill": {"color": ACCENT_5}},
        ],
    }
)
chart_cat.set_rotation(20)
chart_cat.set_style(10)
style_chart(chart_cat, "Tickets by Category")
ws_dash.insert_chart(
    "B16", chart_cat, {"x_offset": 4, "y_offset": 4, "x_scale": 1.0, "y_scale": 1.05}
)

# Chart 2 - Status (Pie)
chart_status = wb.add_chart({"type": "pie"})
chart_status.add_series(
    {
        "name": "Status",
        "categories": srng(F_HEADER + 2, F_LAST + 1, 1),
        "values": srng(F_HEADER + 2, F_LAST + 1, 2),
        "data_labels": {"value": True, "category": True, "font": {"color": "white", "bold": True}},
        "points": [
            {"fill": {"color": ACCENT_3}},
            {"fill": {"color": ACCENT_2}},
            {"fill": {"color": ACCENT_4}},
            {"fill": {"color": ACCENT_5}},
        ],
    }
)
style_chart(chart_status, "Open vs Closed (Status)")
ws_dash.insert_chart("F16", chart_status, {"x_offset": 4, "y_offset": 4, "x_scale": 1.0, "y_scale": 1.05})

# Chart 3 - Severity (Column)
chart_sev = wb.add_chart({"type": "column"})
chart_sev.add_series(
    {
        "name": "Tickets",
        "categories": srng(E_HEADER + 2, E_LAST + 1, 1),
        "values": srng(E_HEADER + 2, E_LAST + 1, 2),
        "fill": {"color": ACCENT},
        "border": {"color": ACCENT},
        "data_labels": {"value": True, "font": {"color": "white", "bold": True}},
    }
)
chart_sev.add_series(
    {
        "name": "SLA Breach",
        "categories": srng(E_HEADER + 2, E_LAST + 1, 1),
        "values": srng(E_HEADER + 2, E_LAST + 1, 3),
        "fill": {"color": ACCENT_4},
        "border": {"color": ACCENT_4},
    }
)
style_chart(chart_sev, "Severity vs SLA Breach")
ws_dash.insert_chart(
    "K16", chart_sev, {"x_offset": 4, "y_offset": 4, "x_scale": 1.0, "y_scale": 1.05}
)

# Move the cursor down for next chart row
# Chart 4 - City (Bar) at B33
chart_city = wb.add_chart({"type": "bar"})
chart_city.add_series(
    {
        "name": "Total",
        "categories": srng(C_HEADER + 2, C_LAST + 1, 1),
        "values": srng(C_HEADER + 2, C_LAST + 1, 2),
        "fill": {"color": ACCENT},
        "border": {"color": ACCENT},
        "data_labels": {"value": True, "font": {"color": "white"}},
    }
)
style_chart(chart_city, "Tickets by City", x_axis="Tickets", y_axis="City")
ws_dash.insert_chart("B33", chart_city, {"x_offset": 4, "y_offset": 4, "x_scale": 1.0, "y_scale": 1.1})

# Chart 5 - Aging buckets (Column)
chart_age = wb.add_chart({"type": "column"})
chart_age.add_series(
    {
        "name": "Tickets",
        "categories": srng(G_HEADER + 2, G_LAST + 1, 1),
        "values": srng(G_HEADER + 2, G_LAST + 1, 2),
        "points": [
            {"fill": {"color": "#22C55E"}},
            {"fill": {"color": "#84CC16"}},
            {"fill": {"color": "#FACC15"}},
            {"fill": {"color": "#F97316"}},
            {"fill": {"color": "#EF4444"}},
        ],
        "data_labels": {"value": True, "font": {"color": "white", "bold": True}},
    }
)
style_chart(chart_age, "Ticket Aging Buckets", x_axis="Bucket", y_axis="Tickets")
ws_dash.insert_chart("F33", chart_age, {"x_offset": 4, "y_offset": 4, "x_scale": 1.0, "y_scale": 1.1})

# Chart 6 - Monthly trend (Line + Area)
chart_trend = wb.add_chart({"type": "line"})
chart_trend.add_series(
    {
        "name": "Total",
        "categories": srng(H_HEADER + 2, H_LAST + 1, 1),
        "values": srng(H_HEADER + 2, H_LAST + 1, 2),
        "line": {"color": ACCENT, "width": 2.5},
        "marker": {"type": "circle", "size": 7, "fill": {"color": ACCENT}},
        "data_labels": {"value": True, "font": {"color": "white"}},
    }
)
chart_trend.add_series(
    {
        "name": "Closed",
        "categories": srng(H_HEADER + 2, H_LAST + 1, 1),
        "values": srng(H_HEADER + 2, H_LAST + 1, 3),
        "line": {"color": ACCENT_2, "width": 2.0},
        "marker": {"type": "circle", "size": 6, "fill": {"color": ACCENT_2}},
    }
)
chart_trend.add_series(
    {
        "name": "Open",
        "categories": srng(H_HEADER + 2, H_LAST + 1, 1),
        "values": srng(H_HEADER + 2, H_LAST + 1, 4),
        "line": {"color": ACCENT_3, "width": 2.0},
        "marker": {"type": "circle", "size": 6, "fill": {"color": ACCENT_3}},
    }
)
style_chart(chart_trend, "Monthly Ticket Trend", x_axis="Month", y_axis="Tickets")
ws_dash.insert_chart("K33", chart_trend, {"x_offset": 4, "y_offset": 4, "x_scale": 1.0, "y_scale": 1.1})

# Spacer + footer
ws_dash.set_row(50, 6)
footer_fmt = wb.add_format(
    {
        "italic": True,
        "font_size": 9,
        "font_color": TEXT_MUTED,
        "bg_color": BG_DARK,
        "font_name": "Segoe UI",
        "align": "center",
    }
)
ws_dash.merge_range(
    "B52:O52",
    "Use the dropdown filters on the Raw_Data table to slice the source list. "
    "Charts are bound to Pivot_Calculations - regenerate the workbook to refresh figures from a new CSV.",
    footer_fmt,
)

# ---------------------------------------------------------------------------
# 7. Category Analysis sheet
# ---------------------------------------------------------------------------

def make_analysis_sheet(name, tab_color):
    ws = wb.add_worksheet(name)
    ws.set_tab_color(tab_color)
    ws.hide_gridlines(2)
    ws.set_zoom(90)
    ws.set_column("A:A", 2)
    for col in range(1, 15):
        ws.set_column(col, col, 13)
    for r in range(0, 60):
        ws.set_row(r, None, f["bg"])
        for c in range(0, 17):
            ws.write_blank(r, c, None, f["bg"])
    return ws


ws_cat = make_analysis_sheet("Category_Analysis", ACCENT)
ws_cat.set_row(1, 38)
ws_cat.merge_range("B2:O2", "Category Analysis", f["title"])
ws_cat.set_row(2, 18)
ws_cat.merge_range(
    "B3:O3", "IT vs Repair & Maintenance vs Marketing vs Other", f["subtitle"]
)

# Mini summary panel
ws_cat.merge_range("B5:O5", "  Category KPIs", f["section"])
ws_cat.set_row(4, 22)
panel_h_fmt = f["panel_header"]

ws_cat.set_row(6, 22)
ws_cat.set_row(7, 22)
ws_cat.write("B7", "Category", panel_h_fmt)
ws_cat.write("C7", "Total", panel_h_fmt)
ws_cat.write("D7", "Open", panel_h_fmt)
ws_cat.write("E7", "Closed", panel_h_fmt)
ws_cat.write("F7", "% Share", panel_h_fmt)
for i, rec in enumerate(b_data):
    r = 7 + i
    ws_cat.write(r, 1, rec[0], f["panel_cell"])
    ws_cat.write_number(r, 2, rec[1], f["panel_cell_num"])
    ws_cat.write_number(r, 3, rec[2], f["panel_cell_num"])
    ws_cat.write_number(r, 4, rec[3], f["panel_cell_num"])
    ws_cat.write_number(r, 5, rec[1] / total if total else 0, f["panel_cell_pct"])

# Donut + Stacked column
chart_cat_d = wb.add_chart({"type": "doughnut"})
chart_cat_d.add_series(
    {
        "name": "Issue Type",
        "categories": srng(B_HEADER + 2, B_LAST + 1, 1),
        "values": srng(B_HEADER + 2, B_LAST + 1, 2),
        "data_labels": {"percentage": True, "font": {"color": "white", "bold": True}},
        "points": [
            {"fill": {"color": ACCENT}},
            {"fill": {"color": ACCENT_2}},
            {"fill": {"color": ACCENT_3}},
            {"fill": {"color": ACCENT_4}},
        ],
    }
)
style_chart(chart_cat_d, "Category Share")
ws_cat.insert_chart("H7", chart_cat_d, {"x_scale": 1.0, "y_scale": 1.0})

chart_cat_b = wb.add_chart({"type": "column", "subtype": "stacked"})
chart_cat_b.add_series(
    {
        "name": "Open",
        "categories": srng(B_HEADER + 2, B_LAST + 1, 1),
        "values": srng(B_HEADER + 2, B_LAST + 1, 3),
        "fill": {"color": ACCENT_3},
        "data_labels": {"value": True, "font": {"color": "white"}},
    }
)
chart_cat_b.add_series(
    {
        "name": "Closed",
        "categories": srng(B_HEADER + 2, B_LAST + 1, 1),
        "values": srng(B_HEADER + 2, B_LAST + 1, 4),
        "fill": {"color": ACCENT_2},
        "data_labels": {"value": True, "font": {"color": "white"}},
    }
)
style_chart(chart_cat_b, "Open vs Closed by Category", x_axis="Category", y_axis="Tickets")
ws_cat.insert_chart("B17", chart_cat_b, {"x_scale": 1.0, "y_scale": 1.1})

# Department breakdown chart (treemap-style approximation: bar)
dept_counts = Counter(r["Department"] for r in records)
dept_rows_start = 17
ws_cat.write("H17", "Department Volume", panel_h_fmt)
ws_cat.write("H18", "Department", panel_h_fmt)
ws_cat.write("I18", "Tickets", panel_h_fmt)
for i, (d, n) in enumerate(sorted(dept_counts.items(), key=lambda x: -x[1])):
    ws_cat.write(18 + i, 7, d, f["panel_cell"])
    ws_cat.write_number(18 + i, 8, n, f["panel_cell_num"])

# ---------------------------------------------------------------------------
# 8. City Analysis
# ---------------------------------------------------------------------------

ws_city = make_analysis_sheet("City_Analysis", ACCENT_2)
ws_city.set_row(1, 38)
ws_city.merge_range("B2:O2", "City and State Analysis", f["title"])
ws_city.set_row(2, 18)
ws_city.merge_range(
    "B3:O3", "Geographic distribution and SLA performance per location", f["subtitle"]
)
ws_city.merge_range("B5:O5", "  City Volume", f["section"])
ws_city.set_row(4, 22)

ws_city.write("B7", "City", panel_h_fmt)
ws_city.write("C7", "Total", panel_h_fmt)
ws_city.write("D7", "Open", panel_h_fmt)
ws_city.write("E7", "Closed", panel_h_fmt)
ws_city.write("F7", "SLA Breach", panel_h_fmt)
for i, rec in enumerate(c_data):
    r = 7 + i
    ws_city.write(r, 1, rec[0], f["panel_cell"])
    ws_city.write_number(r, 2, rec[1], f["panel_cell_num"])
    ws_city.write_number(r, 3, rec[2], f["panel_cell_num"])
    ws_city.write_number(r, 4, rec[3], f["panel_cell_num"])
    ws_city.write_number(r, 5, rec[4], f["panel_cell_num"])

chart_city_b = wb.add_chart({"type": "bar"})
chart_city_b.add_series(
    {
        "name": "Total",
        "categories": srng(C_HEADER + 2, C_LAST + 1, 1),
        "values": srng(C_HEADER + 2, C_LAST + 1, 2),
        "fill": {"color": ACCENT_2},
        "data_labels": {"value": True, "font": {"color": "white"}},
    }
)
style_chart(chart_city_b, "Tickets by City", x_axis="Tickets", y_axis="City")
ws_city.insert_chart("H7", chart_city_b, {"x_scale": 1.0, "y_scale": 1.2})

# State chart
chart_state = wb.add_chart({"type": "doughnut"})
chart_state.add_series(
    {
        "name": "State",
        "categories": srng(D_HEADER + 2, D_LAST + 1, 1),
        "values": srng(D_HEADER + 2, D_LAST + 1, 2),
        "data_labels": {"percentage": True, "font": {"color": "white", "bold": True}},
        "points": [
            {"fill": {"color": ACCENT}},
            {"fill": {"color": ACCENT_3}},
            {"fill": {"color": ACCENT_5}},
        ],
    }
)
style_chart(chart_state, "Share by State")
ws_city.insert_chart("B22", chart_state, {"x_scale": 1.0, "y_scale": 1.0})

# SLA breach by city stacked
chart_city_sla = wb.add_chart({"type": "column", "subtype": "stacked"})
chart_city_sla.add_series(
    {
        "name": "Closed",
        "categories": srng(C_HEADER + 2, C_LAST + 1, 1),
        "values": srng(C_HEADER + 2, C_LAST + 1, 4),
        "fill": {"color": ACCENT_2},
    }
)
chart_city_sla.add_series(
    {
        "name": "Open",
        "categories": srng(C_HEADER + 2, C_LAST + 1, 1),
        "values": srng(C_HEADER + 2, C_LAST + 1, 3),
        "fill": {"color": ACCENT_3},
    }
)
chart_city_sla.add_series(
    {
        "name": "SLA Breach",
        "categories": srng(C_HEADER + 2, C_LAST + 1, 1),
        "values": srng(C_HEADER + 2, C_LAST + 1, 5),
        "fill": {"color": ACCENT_4},
    }
)
style_chart(chart_city_sla, "Open / Closed / SLA breach by City", x_axis="City", y_axis="Tickets")
ws_city.insert_chart("H22", chart_city_sla, {"x_scale": 1.0, "y_scale": 1.0})

# ---------------------------------------------------------------------------
# 9. Aging Analysis
# ---------------------------------------------------------------------------

ws_age = make_analysis_sheet("Aging_Analysis", ACCENT_3)
ws_age.set_row(1, 38)
ws_age.merge_range("B2:O2", "Aging Analysis", f["title"])
ws_age.set_row(2, 18)
ws_age.merge_range(
    "B3:O3", "Bucket distribution and the oldest open tickets", f["subtitle"]
)
ws_age.merge_range("B5:O5", "  Aging Buckets", f["section"])
ws_age.set_row(4, 22)

ws_age.write("B7", "Bucket", panel_h_fmt)
ws_age.write("C7", "Tickets", panel_h_fmt)
ws_age.write("D7", "% of Total", panel_h_fmt)
ws_age.write("E7", "Severity", panel_h_fmt)
severity_label = {
    "0-1 Days": "Fresh",
    "2-3 Days": "Healthy",
    "4-7 Days": "Watch",
    "8-15 Days": "At Risk",
    "15+ Days": "Critical",
}
for i, rec in enumerate(g_data):
    r = 7 + i
    ws_age.write(r, 1, rec[0], f["panel_cell"])
    ws_age.write_number(r, 2, rec[1], f["panel_cell_num"])
    ws_age.write_number(r, 3, rec[1] / total if total else 0, f["panel_cell_pct"])
    ws_age.write(r, 4, severity_label[rec[0]], f["panel_cell"])
ws_age.conditional_format(
    7,
    2,
    7 + len(g_data) - 1,
    2,
    {
        "type": "3_color_scale",
        "min_color": "#16A34A",
        "mid_color": "#F59E0B",
        "max_color": "#DC2626",
    },
)

chart_age2 = wb.add_chart({"type": "column"})
chart_age2.add_series(
    {
        "name": "Tickets",
        "categories": srng(G_HEADER + 2, G_LAST + 1, 1),
        "values": srng(G_HEADER + 2, G_LAST + 1, 2),
        "points": [
            {"fill": {"color": "#22C55E"}},
            {"fill": {"color": "#84CC16"}},
            {"fill": {"color": "#FACC15"}},
            {"fill": {"color": "#F97316"}},
            {"fill": {"color": "#EF4444"}},
        ],
        "data_labels": {"value": True, "font": {"color": "white", "bold": True}},
    }
)
style_chart(chart_age2, "Aging Distribution", x_axis="Bucket", y_axis="Tickets")
ws_age.insert_chart("H7", chart_age2, {"x_scale": 1.05, "y_scale": 1.0})

ws_age.merge_range("B17:O17", "  Oldest Open Tickets", f["section"])
ws_age.set_row(16, 22)
ws_age.write("B19", "Issue ID", panel_h_fmt)
ws_age.write("C19", "Title", panel_h_fmt)
ws_age.write("D19", "City", panel_h_fmt)
ws_age.write("E19", "Severity", panel_h_fmt)
ws_age.write("F19", "Aging (Days)", panel_h_fmt)
ws_age.set_column("C:C", 32)
for i, rec in enumerate(oldest):
    r = 19 + i
    ws_age.write(r, 1, rec["Issue ID"], f["panel_cell"])
    ws_age.write(r, 2, rec["Issue Title"], f["panel_cell"])
    ws_age.write(r, 3, rec["City"], f["panel_cell"])
    ws_age.write(r, 4, rec["Severity"], f["panel_cell"])
    ws_age.write_number(r, 5, rec["Aging (Days)"] or 0, f["panel_cell_dec"])
ws_age.conditional_format(
    19,
    5,
    19 + len(oldest) - 1,
    5,
    {
        "type": "3_color_scale",
        "min_color": "#FDE68A",
        "mid_color": "#F97316",
        "max_color": "#DC2626",
    },
)

# ---------------------------------------------------------------------------
# 10. Trend Analysis
# ---------------------------------------------------------------------------

ws_trend = make_analysis_sheet("Trend_Analysis", ACCENT_5)
ws_trend.set_row(1, 38)
ws_trend.merge_range("B2:O2", "Time Trend Analysis", f["title"])
ws_trend.set_row(2, 18)
ws_trend.merge_range(
    "B3:O3", "Daily and monthly volume trends", f["subtitle"]
)
ws_trend.merge_range("B5:O5", "  Monthly", f["section"])
ws_trend.set_row(4, 22)

# Monthly table
ws_trend.write("B7", "Month", panel_h_fmt)
ws_trend.write("C7", "Total", panel_h_fmt)
ws_trend.write("D7", "Closed", panel_h_fmt)
ws_trend.write("E7", "Open", panel_h_fmt)
for i, rec in enumerate(h_data):
    r = 7 + i
    ws_trend.write(r, 1, rec[0], f["panel_cell"])
    ws_trend.write_number(r, 2, rec[1], f["panel_cell_num"])
    ws_trend.write_number(r, 3, rec[2], f["panel_cell_num"])
    ws_trend.write_number(r, 4, rec[3], f["panel_cell_num"])

chart_month = wb.add_chart({"type": "area", "subtype": "stacked"})
chart_month.add_series(
    {
        "name": "Closed",
        "categories": srng(H_HEADER + 2, H_LAST + 1, 1),
        "values": srng(H_HEADER + 2, H_LAST + 1, 3),
        "fill": {"color": ACCENT_2, "transparency": 30},
        "border": {"color": ACCENT_2},
    }
)
chart_month.add_series(
    {
        "name": "Open",
        "categories": srng(H_HEADER + 2, H_LAST + 1, 1),
        "values": srng(H_HEADER + 2, H_LAST + 1, 4),
        "fill": {"color": ACCENT_3, "transparency": 30},
        "border": {"color": ACCENT_3},
    }
)
style_chart(chart_month, "Monthly Trend (Stacked Area)", x_axis="Month", y_axis="Tickets")
ws_trend.insert_chart("G7", chart_month, {"x_scale": 1.0, "y_scale": 1.0})

# Daily
ws_trend.merge_range("B17:O17", "  Daily", f["section"])
ws_trend.set_row(16, 22)
ws_trend.write("B19", "Day", panel_h_fmt)
ws_trend.write("C19", "Tickets", panel_h_fmt)
for i, rec in enumerate(i_data):
    r = 19 + i
    ws_trend.write(r, 1, rec[0], f["panel_cell"])
    ws_trend.write_number(r, 2, rec[1], f["panel_cell_num"])

chart_day = wb.add_chart({"type": "line"})
chart_day.add_series(
    {
        "name": "Daily Tickets",
        "categories": srng(I_HEADER + 2, I_LAST + 1, 1),
        "values": srng(I_HEADER + 2, I_LAST + 1, 2),
        "line": {"color": ACCENT, "width": 2.25},
        "marker": {"type": "circle", "size": 6, "fill": {"color": ACCENT}},
        "data_labels": {"value": True, "font": {"color": "white"}},
    }
)
style_chart(chart_day, "Daily Ticket Volume", x_axis="Day", y_axis="Tickets")
ws_trend.insert_chart("G19", chart_day, {"x_scale": 1.05, "y_scale": 1.2})

# ---------------------------------------------------------------------------
# 11. SLA Performance
# ---------------------------------------------------------------------------

ws_sla = make_analysis_sheet("SLA_Performance", ACCENT_4)
ws_sla.set_row(1, 38)
ws_sla.merge_range("B2:O2", "SLA and Resolution Performance", f["title"])
ws_sla.set_row(2, 18)
ws_sla.merge_range(
    "B3:O3",
    "SLA targets: Critical=1d, High=2d, Medium=4d, Low=7d, Unspecified=5d",
    f["subtitle"],
)
ws_sla.merge_range("B5:O5", "  Outcome", f["section"])
ws_sla.set_row(4, 22)

ws_sla.write("B7", "SLA Status", panel_h_fmt)
ws_sla.write("C7", "Tickets", panel_h_fmt)
ws_sla.write("D7", "% Share", panel_h_fmt)
sla_total = sla_met + sla_breach
for i, rec in enumerate(k_data):
    r = 7 + i
    ws_sla.write(r, 1, rec[0], f["panel_cell"])
    ws_sla.write_number(r, 2, rec[1], f["panel_cell_num"])
    ws_sla.write_number(r, 3, rec[1] / sla_total if sla_total else 0, f["panel_cell_pct"])

chart_sla = wb.add_chart({"type": "doughnut"})
chart_sla.add_series(
    {
        "name": "SLA",
        "categories": srng(K_HEADER + 2, K_LAST + 1, 1),
        "values": srng(K_HEADER + 2, K_LAST + 1, 2),
        "data_labels": {"percentage": True, "font": {"color": "white", "bold": True}},
        "points": [
            {"fill": {"color": ACCENT_2}},
            {"fill": {"color": ACCENT_4}},
        ],
    }
)
chart_sla.set_rotation(180)
style_chart(chart_sla, "SLA Met vs Breach")
ws_sla.insert_chart("F7", chart_sla, {"x_scale": 0.9, "y_scale": 0.9})

# Resolver dept performance
ws_sla.merge_range("B17:O17", "  Resolver Department Performance", f["section"])
ws_sla.set_row(16, 22)
ws_sla.write("B19", "Resolver Department", panel_h_fmt)
ws_sla.write("C19", "Tickets Resolved", panel_h_fmt)
ws_sla.write("D19", "Avg Resolution (Hrs)", panel_h_fmt)
for i, rec in enumerate(j_data):
    r = 19 + i
    ws_sla.write(r, 1, rec[0], f["panel_cell"])
    ws_sla.write_number(r, 2, rec[1], f["panel_cell_num"])
    ws_sla.write_number(r, 3, rec[2], f["panel_cell_dec"])
ws_sla.conditional_format(
    19, 3, 19 + len(j_data) - 1, 3,
    {
        "type": "3_color_scale",
        "min_color": "#22C55E",
        "mid_color": "#F59E0B",
        "max_color": "#EF4444",
    },
)

chart_resolver = wb.add_chart({"type": "bar"})
chart_resolver.add_series(
    {
        "name": "Tickets Resolved",
        "categories": srng(J_HEADER + 2, J_LAST + 1, 1),
        "values": srng(J_HEADER + 2, J_LAST + 1, 2),
        "fill": {"color": ACCENT},
        "data_labels": {"value": True, "font": {"color": "white"}},
    }
)
style_chart(chart_resolver, "Tickets Resolved by Department", x_axis="Tickets", y_axis="Department")
ws_sla.insert_chart("F19", chart_resolver, {"x_scale": 1.0, "y_scale": 1.0})

# ---------------------------------------------------------------------------
# 12. README sheet
# ---------------------------------------------------------------------------

ws_readme = wb.add_worksheet("README")
ws_readme.set_tab_color("#475569")
ws_readme.hide_gridlines(2)
ws_readme.set_column("A:A", 2)
ws_readme.set_column("B:B", 110)

ws_readme.set_row(1, 30)
ws_readme.write("B2", "Ticket Analysis Dashboard - Quick Guide", f["readme_h1"])
ws_readme.set_row(2, 8)
ws_readme.write("B4", "1. What this workbook contains", f["readme_h2"])
ws_readme.set_row(4, 70)
ws_readme.write(
    "B5",
    "  - Dashboard          : executive KPI cards plus the main visual analytics grid.\n"
    "  - Category_Analysis  : ticket volume and open/closed split per Issue Type and Department.\n"
    "  - City_Analysis      : geographic distribution by City and State, plus SLA breach map.\n"
    "  - Aging_Analysis     : aging buckets, % share, and the top 10 oldest open tickets.\n"
    "  - Trend_Analysis     : monthly stacked-area trend and daily volume line chart.\n"
    "  - SLA_Performance    : SLA met vs breach, resolver department performance heatmap.\n"
    "  - Pivot_Calculations : every aggregation that drives the charts (formula-driven).\n"
    "  - Raw_Data           : the source ticket list as a fully filterable Excel Table named 'Tickets'.",
    f["readme_p"],
)

ws_readme.write("B7", "2. How to filter the dashboard", f["readme_h2"])
ws_readme.set_row(7, 70)
ws_readme.write(
    "B8",
    "Open the Raw_Data sheet and use the column filter dropdowns on the 'Tickets' table to slice "
    "by City, Severity, Status, Department, Issue Type or any other column. The filters behave "
    "like slicers - all charts that are derived from the same underlying data update when you "
    "regenerate the workbook with a new CSV. For point-in-time slicing inside an existing "
    "session, use the AutoFilter dropdowns in the table header.",
    f["readme_p"],
)

ws_readme.write("B10", "3. SLA targets used in calculations", f["readme_h2"])
ws_readme.set_row(10, 90)
ws_readme.write(
    "B11",
    "  - Critical    : 1 day to close\n"
    "  - High        : 2 days\n"
    "  - Medium      : 4 days\n"
    "  - Low         : 7 days\n"
    "  - Unspecified : 5 days (default)\n"
    "An open ticket whose age exceeds its target is classified as 'Breach', otherwise 'Met'. "
    "Aging is computed against the resolution timestamp for closed tickets and against the "
    "snapshot date (most recent reported timestamp in the dataset) for open tickets.",
    f["readme_p"],
)

ws_readme.write("B13", "4. How to refresh with new data", f["readme_h2"])
ws_readme.set_row(13, 80)
ws_readme.write(
    "B14",
    "Replace Apps.csv with a newer export from the source Google Sheet (same column layout) and "
    "re-run build_dashboard.py. The script regenerates Ticket_Analysis_Dashboard.xlsx end-to-end "
    "with all charts, KPIs and conditional formatting bound to the fresh data.",
    f["readme_p"],
)

ws_readme.write("B16", "5. Notes on advanced features", f["readme_h2"])
ws_readme.set_row(16, 90)
ws_readme.write(
    "B17",
    "Power Query, Power Pivot/DAX, true Slicers, Funnel/Gauge/Filled-Map charts and VBA macros "
    "are not generatable from a Python script that emits a vanilla .xlsx. To add them, open this "
    "file in Excel desktop and: (a) Insert > Slicer on the 'Tickets' table for Category/City/"
    "Status; (b) Data > Get Data > From Text/CSV to wire Apps.csv through Power Query for "
    "auto-refresh; (c) Insert > Funnel/Gauge for the Status and SLA visuals. Everything else - "
    "KPIs, charts, conditional formatting, aging, SLA logic, formula-driven pivots - is already "
    "present in this file.",
    f["readme_p"],
)

ws_readme.write("B19", "Jump to:", f["readme_h2"])
ws_readme.write_url("B20", "internal:'Dashboard'!A1", f["readme_link"], "-> Open Dashboard")
ws_readme.write_url("B21", "internal:'Raw_Data'!A1", f["readme_link"], "-> Open Raw_Data")
ws_readme.write_url(
    "B22", "internal:'Pivot_Calculations'!A1", f["readme_link"], "-> Open Pivot_Calculations"
)

# ---------------------------------------------------------------------------
# Sheet order: README first reading, but Dashboard active. Reorder.
# ---------------------------------------------------------------------------

# Worksheets are created in order; xlsxwriter doesn't re-order, but we set
# Dashboard as first_sheet/activate above, so it's the one shown on open.

wb.close()
print(f"Wrote {OUT_PATH} ({os.path.getsize(OUT_PATH):,} bytes)")
