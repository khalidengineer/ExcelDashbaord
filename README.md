# Ticket Intelligence & Operations Monitoring Platform (TIOMP)

> A premium, NOC-grade Excel VBA dashboard built end-to-end from `Apps.csv`.
> Feels like ServiceNow Analytics × Splunk Operations Center × Power BI
> Executive Center — but it lives entirely inside one `.xlsm` file.

![banner](https://img.shields.io/badge/platform-Excel%20VBA-1a2332?style=flat-square)
![status](https://img.shields.io/badge/status-production--ready-10B981?style=flat-square)
![version](https://img.shields.io/badge/version-1.0.0-00E5FF?style=flat-square)

---

## What is this?

A complete, modular, enterprise-grade IT Operations analytics platform that
imports the `Apps.csv` ticket dataset and produces:

- 16 sheets (Dashboard, 8 specialized modules, 5 data sheets, Settings, Logs, Hidden Config)
- 18 enterprise KPIs + 36 underlying metric values
- AI-style analytics: alerts, recommendations, breach predictions, root-cause clustering
- 13+ dark-themed charts in NOC aesthetic with neon cyan highlights
- Floating navigation rail, smart search, drill-through, PDF export, email MIS, role-based login
- One-click build via `BuildEnterpriseDashboard` macro

**Single dependency:** `Apps.csv` next to the workbook. Nothing else.

---

## Quick Start

```
1. Open Excel, save a new workbook as TIOMP.xlsm next to Apps.csv
2. Press Alt+F11 -> File -> Import File -> import all 7 .bas modules from /vba
3. Paste the ThisWorkbook.cls code into the existing ThisWorkbook object
4. Press Alt+F8 -> run BuildEnterpriseDashboard
```

See [DEPLOYMENT.md](DEPLOYMENT.md) for the full step-by-step.

---

## Architecture

```
                     +----------------------------+
                     |   M01_Builder              |
                     |   * BuildEnterpriseDashboard
                     |   * RefreshAll             |
                     |   * Theme + Logger         |
                     +-------------+--------------+
                                   |
   +---------------------+---------+---------+----------------------+
   |                     |                   |                      |
   v                     v                   v                      v
+-----------+    +---------------+    +-------------+      +----------------+
| M02_Data  |    | M03_KpiEngine |    | M04_Dashbrd |      | M07_Interaction|
| Engine    |    | + AI Layer    |    | UI          |      | Nav, search,   |
| CSV->Clean|    | KPIs, alerts, |    | KPI cards,  |      | drill, exports,|
| ->Model   |    | recos, preds  |    | header, feed|      | login, theme   |
+-----------+    +---------------+    +-------------+      +----------------+
                                   |
                                   v
                          +-----------------+
                          | M05_Charts      |
                          | dark-themed     |
                          | chart factory   |
                          +-----------------+
                                   |
                                   v
                       +----------------------+
                       | M06_Modules          |
                       | 8 analytics sheets   |
                       +----------------------+
```

| Module                | Responsibility                                            |
| --------------------- | --------------------------------------------------------- |
| **M01_Builder**       | Single entry point, theme constants, logger, perf helpers |
| **M02_DataEngine**    | CSV parser, normalizer, derived columns, aggregations     |
| **M03_KpiEngine**     | 36 KPIs, AI alerts, recommendations, breach prediction    |
| **M04_DashboardUI**   | Premium dashboard shell, KPI cards, AI feed, console      |
| **M05_Charts**        | Dark-themed chart factory (line, donut, gauge, radar, etc)|
| **M06_Modules**       | 8 specialized analytics sheets                            |
| **M07_Interaction**   | Navigation, search, drill, exports, login, theme, archive |
| **ThisWorkbook**      | Workbook event handlers (open/close/sheet activate)       |

---

## Sheets Produced

| Sheet                  | Purpose                                                   |
| ---------------------- | --------------------------------------------------------- |
| `Dashboard_Main`       | Executive Command Center — hero KPIs, charts, AI feed     |
| `Executive_View`       | Board-level radar, gauge, trend                           |
| `SLA_Intelligence`     | Compliance gauge, breach prediction table                 |
| `Incident_Analytics`   | War room — active critical/high tickets                   |
| `Agent_Analytics`      | Resolver leaderboard with Star/Solid/Watch/Coach tiers    |
| `Application_Analytics`| Issue type, location and category health                  |
| `RCA_Analytics`        | Auto-clustered root-cause tree + AI recommendations       |
| `Forecast_Analytics`   | 3-day forecast, peak hour pill, hour×DOW heatmap          |
| `Risk_Analytics`       | Risk matrix with color scale + ops radar                  |
| `Settings`             | SLA tuning, theme, security, action buttons               |
| `Logs`                 | Structured activity log                                   |
| `Raw_Data`*            | Imported CSV (hidden)                                     |
| `Clean_Data`*          | Normalized + 21 derived columns (hidden)                  |
| `Data_Model`*          | Aggregations / lookups (hidden)                           |
| `KPI_Engine`*          | KPI store + AI layer outputs (hidden)                     |
| `Hidden_Config`*       | Theme, credentials, recipients (very hidden)              |

\* = hidden by default for a polished UX

---

## KPIs Computed

**Volume & throughput:** Total Tickets, Open, Closed, Closure Rate, Backlog Index, Trend Velocity (W/W), Auto Closure Rate
**Severity:** Critical Total, Critical Open, High, Medium, Low, Unspecified
**SLA:** SLA Compliance %, Breached, At Risk, MTTR, Acknowledgement Time
**Health:** Stability Score, Operational Health, Customer Satisfaction Index, Agent Efficiency, Support Utilization
**Risk:** Risk Score, Downtime Risk, Aging Risk, Priority Heat, Critical Failure Rate, Impact Score, Repeat Incident %, Escalation Rate
**Capacity:** Aging > 3 days, Aging > 7 days, Peak Hour

---

## AI-Style Analytics

| Feature                   | Implementation                                            |
| ------------------------- | --------------------------------------------------------- |
| Smart Alert Generation    | Threshold rules across SLA, criticals, repeats, aging     |
| AI Recommendations        | 6 themes: failure cluster, reporter, resolver, off-hours, hygiene, hot-spot |
| Breach Prediction         | Open tickets ranked by `age / SLA target`                |
| Root-Cause Clustering     | Keyword-based category extraction (CCTV, Refrigeration, IT, Electrical, etc.) |
| Repeat Detection          | Same reporter + similar title key                         |
| Severity Classifier       | Keyword-based suggestion for blanks (`fire`→Critical, `not working`→High …) |
| Peak Hour Detection       | Argmax over hour-of-day distribution                      |
| Trend Forecast            | 3-day moving-average projection with slope                |
| High-Risk Site Detection  | Avg risk × volume per location                            |

---

## Visuals

- Animated KPI Cards with neon-cyan accents and glyph badges
- Operational Pulse trend chart with 3-day forecast
- Severity donut + Issue Type pie
- SLA semi-doughnut Speedometer Gauge
- Aging Funnel with bucket-graded color
- Top Locations / Top Resolvers horizontal bars
- Resolver Speed-vs-SLA scatter
- Operational Health radar (6 dimensions)
- Severity × Status incident risk matrix (3-color scale)
- Hour × Day-of-Week heatmap
- AI Alert Feed (colored cards stacked)
- Live Activity Console (latest 8 tickets, NOC log style)
- Root-Cause Tree with shape connectors

---

## Premium UI System

| Element        | Specification                                              |
| -------------- | ---------------------------------------------------------- |
| Background     | `#0F171A` near-black canvas                                |
| Panels         | `#1A2332` with subtle drop-shadow (msoShadow21, 8px blur)  |
| Accent         | `#00E5FF` neon cyan (titles, active nav, primary action)   |
| Secondary      | `#00B3FF` blue, `#9B54F6` purple, `#EC60A6` pink           |
| Status         | `#10B981` good, `#F59E0B` warn, `#EF4444` critical         |
| Typography     | Segoe UI Semibold (titles), Segoe UI (body), Consolas (logs)|
| Cards          | Rounded rectangles, 8% corner radius, gradient bar         |
| Iconography    | Unicode symbol glyphs (no external dependencies)           |
| Effects        | Drop shadows, accent strips, neon active states            |

---

## Performance

- Single-pass array reads/writes (no per-cell loops on hot paths)
- `Scripting.Dictionary` for all aggregations
- `ScreenUpdating`, `EnableEvents`, `Calculation` toggled during builds
- Charts redraw via lightweight `RefreshDashboardMain` not full rebuild
- Typical full-build time: **2–5 seconds** on 60 rows; **<10s** on 5000 rows

---

## License

This codebase is provided as-is for the repository owner's use. Adapt freely
for internal enterprise reporting needs.
