Attribute VB_Name = "TIOMP_UI"
Option Explicit
Option Compare Text

' === TIOMP_UI - Part of TIOMP Dashboard ===
' Imports: M04_DashboardUI, M05_Charts

'==============================================================================
' MODULE-LEVEL DECLARATIONS (must be at top)
'==============================================================================


' --- declarations from M04_DashboardUI ---
'==============================================================================
' MODULE      : M04_DashboardUI
' DESCRIPTION : Builds the executive Dashboard_Main page with premium dark NOC
'               aesthetic: header bar, side navigation, KPI cards, AI feed,
'               activity console and chart hosts. Uses Excel shapes for visual
'               primitives (cards, panels, neon accents) and embedded charts.
'==============================================================================

' Dashboard layout grid (cells)
Private Const G_TITLE_ROW   As Long = 1   ' header
Private Const G_KPI_ROW     As Long = 4   ' kpi cards row anchor
Private Const G_BODY_ROW    As Long = 11  ' body content begins
Private Const G_NAV_COL     As Long = 1
Private Const G_BODY_COL    As Long = 2

'==============================================================================
'                          PUBLIC ENTRY POINTS
'==============================================================================

'==============================================================================
' PROCEDURES
'==============================================================================

'==============================================================================
' === SECTION: M04_DashboardUI ===
'==============================================================================
Public Sub BuildDashboardMain()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_DASH)
    ResetSheetForDashboard ws

    ' 1. Sheet canvas - dark background
    PaintCanvas ws

    ' 2. Top header bar
    BuildHeaderBar ws, "EXECUTIVE COMMAND CENTER", _
        "Real-time Ticket Intelligence  |  AI-Augmented Operations"

    ' 3. Side navigation (left rail)
    BuildNavigationRail ws, SHT_DASH

    ' 4. Hero KPI cards row (8 primary KPIs)
    BuildPrimaryKpiCards ws

    ' 5. Secondary KPI strip (10 supporting KPIs)
    BuildSecondaryKpiStrip ws

    ' 6. Center content - charts cluster
    BuildDashboardCharts ws

    ' 7. AI alert feed and activity console
    BuildAiAlertPanel ws
    BuildActivityConsole ws

    ' 8. Footer bar with refresh button + version
    BuildFooterBar ws

    ' Final touches
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Range("A1").Select
    LogInfo "UI", "Dashboard_Main built"
End Sub

Public Sub RefreshDashboardMain()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_DASH)
    On Error GoTo 0
    If ws Is Nothing Then
        ' Dashboard wasn't built yet - do a full build first
        BuildEnterpriseDashboard
        Exit Sub
    End If
    ' Repaint KPI cards by clearing+rebuilding shapes prefixed kpi_*
    DeleteShapesByPrefix ws, "kpi_"
    DeleteShapesByPrefix ws, "ai_"
    DeleteShapesByPrefix ws, "act_"
    BuildPrimaryKpiCards ws
    BuildSecondaryKpiStrip ws
    BuildAiAlertPanel ws
    BuildActivityConsole ws
    RefreshAllCharts
End Sub

'==============================================================================
'                          CANVAS / RESET
'==============================================================================
Public Sub ResetSheetForDashboard(ws As Worksheet)
    Dim sh As Shape
    Application.DisplayAlerts = False
    For Each sh In ws.Shapes
        sh.Delete
    Next sh
    ws.Cells.Clear
    On Error Resume Next
    ws.ChartObjects.Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Public Sub PaintCanvas(ws As Worksheet)
    ' Tighten column widths to a consistent grid (28 cols × ~22 rows visible viewport)
    Dim c As Long
    For c = 1 To 28
        ws.Columns(c).ColumnWidth = 9.5
    Next c
    ws.Rows.RowHeight = 16
    ws.Cells.Interior.Color = CLR_BG
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Color = CLR_TEXT
    ws.Cells.Font.Size = 9
    ' Pad nav column
    ws.Columns(1).ColumnWidth = 22
End Sub

'==============================================================================
'                          HEADER BAR
'==============================================================================
Public Sub BuildHeaderBar(ws As Worksheet, ByVal title As String, ByVal subtitle As String)
    Dim hdr As Shape, accent As Shape, dot As Shape, brand As Shape
    Dim left As Double, top As Double, width As Double

    left = ws.Range("A1").Left
    top = ws.Range("A1").Top
    width = ws.Range("A1:AB1").Width

    Set hdr = AddShape(ws, msoShapeRectangle, "hdr_bar", left, top, width, HEADER_HEIGHT)
    StyleShape hdr, CLR_PANEL, 0
    AddShadow hdr

    ' Neon cyan accent line at bottom of header
    Set accent = AddShape(ws, msoShapeRectangle, "hdr_accent", left, top + HEADER_HEIGHT - 3, width, 3)
    StyleShape accent, CLR_ACCENT, 0

    ' Brand glyph (rounded square with logo letter)
    Set brand = AddShape(ws, msoShapeRoundedRectangle, "hdr_brand", left + 14, top + 14, 42, 42)
    StyleShape brand, CLR_ACCENT, 0
    brand.TextFrame2.TextRange.Text = "TI"
    With brand.TextFrame2.TextRange.Font
        .Bold = msoTrue: .Size = 18: .Name = "Segoe UI Black"
        .Fill.ForeColor.RGB = CLR_BG
    End With
    brand.TextFrame2.HorizontalAnchor = msoAnchorCenter
    brand.TextFrame2.VerticalAnchor = msoAnchorMiddle

    ' Title text
    Dim t As Shape
    Set t = AddShape(ws, msoShapeRectangle, "hdr_title", left + 70, top + 8, 600, 30)
    StyleShape t, CLR_PANEL, 0
    t.TextFrame2.TextRange.Text = title
    With t.TextFrame2.TextRange.Font
        .Bold = msoTrue: .Size = 16: .Name = "Segoe UI Semibold"
        .Fill.ForeColor.RGB = CLR_TEXT
    End With
    t.TextFrame2.HorizontalAnchor = msoAnchorCenter
    t.TextFrame2.VerticalAnchor = msoAnchorMiddle
    t.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

    Dim s As Shape
    Set s = AddShape(ws, msoShapeRectangle, "hdr_sub", left + 70, top + 38, 600, 22)
    StyleShape s, CLR_PANEL, 0
    s.TextFrame2.TextRange.Text = subtitle
    With s.TextFrame2.TextRange.Font
        .Size = 10: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_MUTED
    End With
    s.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    s.TextFrame2.VerticalAnchor = msoAnchorMiddle

    ' Right-aligned status pill: live indicator
    Dim pill As Shape
    Set pill = AddShape(ws, msoShapeRoundedRectangle, "hdr_pill", left + width - 220, top + 18, 200, 34)
    StyleShape pill, CLR_PANEL_HI, 0
    pill.TextFrame2.TextRange.Text = ChrW(9679) & "  LIVE  |  " & Format(Now, "dd-mmm hh:mm")
    With pill.TextFrame2.TextRange.Font
        .Bold = msoTrue: .Size = 10: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_GOOD
    End With
    pill.TextFrame2.HorizontalAnchor = msoAnchorCenter
    pill.TextFrame2.VerticalAnchor = msoAnchorMiddle
End Sub

'==============================================================================
'                          PRIMARY KPI CARDS (Hero row)
' 8 cards in a row across columns B..AB, row index G_KPI_ROW
'==============================================================================
Private Sub BuildPrimaryKpiCards(ws As Worksheet)
    Dim cards As Variant
    cards = Array( _
        Array("Total Tickets", "TotalTickets", "0", CLR_ACCENT, ChrW(8505)), _
        Array("Open", "OpenTickets", "0", CLR_WARN, ChrW(9888)), _
        Array("Critical Open", "CriticalOpen", "0", CLR_BAD, ChrW(9760)), _
        Array("SLA Compliance", "SLA_Compliance", "0.0%", CLR_GOOD, ChrW(10003)), _
        Array("MTTR (h)", "MTTR_Hours", "0.0", CLR_ACCENT2, ChrW(8987)), _
        Array("Backlog Index", "Backlog_Index", "0.0%", CLR_PURPLE, ChrW(9776)), _
        Array("Risk Score", "Risk_Score", "0", CLR_PINK, ChrW(9889)), _
        Array("Ops Health", "OpsHealth_Score", "0", CLR_GOOD, ChrW(9829)))

    Dim startLeft As Double, top As Double, gap As Double
    startLeft = ws.Cells(G_KPI_ROW, 2).Left
    top = ws.Cells(G_KPI_ROW, 2).Top
    gap = 8
    Dim totalW As Double
    totalW = ws.Range("B" & G_KPI_ROW & ":AB" & G_KPI_ROW).Width
    Dim cardW As Double
    cardW = (totalW - gap * 7) / 8

    Dim i As Long
    For i = 0 To 7
        Dim arr As Variant: arr = cards(i)
        DrawKpiCard ws, _
            "kpi_p_" & i, _
            startLeft + i * (cardW + gap), _
            top, cardW, KPI_CARD_H, _
            CStr(arr(0)), CStr(arr(1)), CStr(arr(2)), CLng(arr(3)), CStr(arr(4))
    Next i
End Sub

'==============================================================================
'                       SECONDARY KPI STRIP
' 10 compact KPI strip, row 8, similar look but shorter
'==============================================================================
Private Sub BuildSecondaryKpiStrip(ws As Worksheet)
    Dim cards As Variant
    cards = Array( _
        Array("Closure Rate", "ClosureRate", "0.0%"), _
        Array("Repeat %", "Repeat_Rate", "0.0%"), _
        Array("Escalation %", "Escalation_Rate", "0.0%"), _
        Array("Stability", "Stability_Score", "0"), _
        Array("Impact", "Impact_Score", "0.0"), _
        Array("Aging Risk", "Aging_Risk", "0"), _
        Array("Priority Heat", "Priority_Heat", "0.0"), _
        Array("Trend % WoW", "Trend_Velocity", "0.0"), _
        Array("CSAT Index", "Customer_Satisfaction", "0"), _
        Array("Agent Eff", "Agent_Efficiency", "0"))

    Dim row As Long: row = 8
    Dim startLeft As Double, top As Double, gap As Double
    startLeft = ws.Cells(row, 2).Left
    top = ws.Cells(row, 2).Top
    gap = 6
    Dim totalW As Double
    totalW = ws.Range("B" & row & ":AB" & row).Width
    Dim cardW As Double
    cardW = (totalW - gap * 9) / 10

    Dim i As Long
    For i = 0 To 9
        Dim arr As Variant: arr = cards(i)
        DrawSlimKpi ws, _
            "kpi_s_" & i, _
            startLeft + i * (cardW + gap), _
            top, cardW, 50, _
            CStr(arr(0)), CStr(arr(1)), CStr(arr(2))
    Next i
End Sub

'==============================================================================
'                          KPI CARD DRAWERS
'==============================================================================
Public Sub DrawKpiCard(ws As Worksheet, ByVal name As String, _
        ByVal lft As Double, ByVal tp As Double, _
        ByVal w As Double, ByVal h As Double, _
        ByVal label As String, ByVal kpiKey As String, _
        ByVal numFmt As String, ByVal accent As Long, ByVal glyph As String)

    Dim card As Shape, bar As Shape, lbl As Shape, val As Shape, ic As Shape

    ' Shadowed rounded card
    Set card = AddShape(ws, msoShapeRoundedRectangle, name, lft, tp, w, h)
    card.Adjustments.Item(1) = 0.08
    StyleShape card, CLR_PANEL, 0
    AddShadow card

    ' Left accent strip
    Set bar = AddShape(ws, msoShapeRectangle, name & "_bar", lft + 6, tp + 12, 4, h - 24)
    StyleShape bar, accent, 0

    ' Glyph icon (top-right)
    Set ic = AddShape(ws, msoShapeOval, name & "_ic", lft + w - 38, tp + 10, 28, 28)
    StyleShape ic, accent, 0
    With ic.TextFrame2
        .TextRange.Text = glyph
        With .TextRange.Font
            .Size = 14: .Bold = msoTrue: .Name = "Segoe UI Symbol"
            .Fill.ForeColor.RGB = CLR_BG
        End With
        .HorizontalAnchor = msoAnchorCenter
        .VerticalAnchor = msoAnchorMiddle
    End With

    ' Big value
    Set val = AddShape(ws, msoShapeRectangle, name & "_v", lft + 18, tp + 12, w - 60, 44)
    StyleShape val, CLR_PANEL, 0
    val.TextFrame2.TextRange.Text = FmtKpi(kpiKey, numFmt)
    With val.TextFrame2.TextRange.Font
        .Size = 26: .Bold = msoTrue: .Name = "Segoe UI Semibold"
        .Fill.ForeColor.RGB = CLR_TEXT
    End With
    val.TextFrame2.HorizontalAnchor = msoAnchorCenter
    val.TextFrame2.VerticalAnchor = msoAnchorMiddle
    val.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

    ' Label
    Set lbl = AddShape(ws, msoShapeRectangle, name & "_l", lft + 18, tp + h - 28, w - 24, 20)
    StyleShape lbl, CLR_PANEL, 0
    lbl.TextFrame2.TextRange.Text = UCase$(label)
    With lbl.TextFrame2.TextRange.Font
        .Size = 9: .Bold = msoTrue: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_MUTED
    End With
    lbl.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    lbl.TextFrame2.VerticalAnchor = msoAnchorMiddle
End Sub

Public Sub DrawSlimKpi(ws As Worksheet, ByVal name As String, _
        ByVal lft As Double, ByVal tp As Double, _
        ByVal w As Double, ByVal h As Double, _
        ByVal label As String, ByVal kpiKey As String, ByVal numFmt As String)

    Dim card As Shape, lbl As Shape, val As Shape
    Set card = AddShape(ws, msoShapeRoundedRectangle, name, lft, tp, w, h)
    card.Adjustments.Item(1) = 0.18
    StyleShape card, CLR_PANEL_HI, 0

    Set lbl = AddShape(ws, msoShapeRectangle, name & "_l", lft + 8, tp + 4, w - 16, 14)
    StyleShape lbl, CLR_PANEL_HI, 0
    lbl.TextFrame2.TextRange.Text = UCase$(label)
    With lbl.TextFrame2.TextRange.Font
        .Size = 8: .Bold = msoTrue: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_MUTED
    End With
    lbl.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

    Set val = AddShape(ws, msoShapeRectangle, name & "_v", lft + 8, tp + 18, w - 16, h - 22)
    StyleShape val, CLR_PANEL_HI, 0
    val.TextFrame2.TextRange.Text = FmtKpi(kpiKey, numFmt)
    With val.TextFrame2.TextRange.Font
        .Size = 16: .Bold = msoTrue: .Name = "Segoe UI Semibold"
        .Fill.ForeColor.RGB = CLR_ACCENT
    End With
    val.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    val.TextFrame2.VerticalAnchor = msoAnchorMiddle
End Sub

'==============================================================================
'                          DASHBOARD CHARTS CLUSTER
' Lays out 4 charts in two rows: trend, severity, status, aging
'==============================================================================
Private Sub BuildDashboardCharts(ws As Worksheet)
    ' Section header
    Dim hdr As Shape
    Set hdr = AddShape(ws, msoShapeRectangle, "sec_hdr_charts", _
        ws.Cells(G_BODY_ROW, 2).Left, ws.Cells(G_BODY_ROW, 2).Top, _
        ws.Range("B" & G_BODY_ROW & ":Q" & G_BODY_ROW).Width, 24)
    StyleShape hdr, CLR_BG, 0
    hdr.TextFrame2.TextRange.Text = "OPERATIONS PULSE"
    With hdr.TextFrame2.TextRange.Font
        .Bold = msoTrue: .Size = 11: .Name = "Segoe UI Semibold"
        .Fill.ForeColor.RGB = CLR_ACCENT
    End With
    hdr.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

    Dim chartTop As Double, chartLeft As Double, chartW As Double, chartH As Double, gap As Double
    chartLeft = ws.Cells(G_BODY_ROW + 1, 2).Left
    chartTop = ws.Cells(G_BODY_ROW + 1, 2).Top
    gap = 8
    chartW = (ws.Range("B:Q").Width - gap) / 2
    chartH = 200

    BuildTrendChart ws, chartLeft, chartTop, chartW, chartH, "Daily Ticket Volume + Forecast"
    BuildSeverityDonut ws, chartLeft + chartW + gap, chartTop, chartW, chartH, "Severity Distribution"
    BuildAgingFunnel ws, chartLeft, chartTop + chartH + gap, chartW, chartH, "Ticket Aging Buckets"
    BuildCategoryBars ws, chartLeft + chartW + gap, chartTop + chartH + gap, chartW, chartH, "Category Heat (Auto-Clustered)"
End Sub

'==============================================================================
'                          AI ALERT FEED PANEL (right column)
'==============================================================================
Private Sub BuildAiAlertPanel(ws As Worksheet)
    Dim wsKpi As Worksheet
    Set wsKpi = ThisWorkbook.Worksheets(SHT_KPI)

    Dim left As Double, top As Double, w As Double, h As Double
    left = ws.Cells(G_BODY_ROW, 18).Left
    top = ws.Cells(G_BODY_ROW, 18).Top
    w = ws.Range("R:AB").Width
    h = 230

    ' Section header
    Dim hdr As Shape
    Set hdr = AddShape(ws, msoShapeRectangle, "ai_hdr", left, top, w, 24)
    StyleShape hdr, CLR_BG, 0
    hdr.TextFrame2.TextRange.Text = "AI ALERT FEED  " & ChrW(9889)
    With hdr.TextFrame2.TextRange.Font
        .Bold = msoTrue: .Size = 11: .Name = "Segoe UI Semibold"
        .Fill.ForeColor.RGB = CLR_ACCENT
    End With

    ' Panel
    Dim pnl As Shape
    Set pnl = AddShape(ws, msoShapeRoundedRectangle, "ai_pnl", left, top + 28, w, h)
    pnl.Adjustments.Item(1) = 0.04
    StyleShape pnl, CLR_PANEL, 0
    AddShadow pnl

    ' Read alerts from KPI sheet (col D..G under "AI ALERTS" header)
    Dim alertRows As Collection: Set alertRows = New Collection
    Dim r As Long, lr As Long, foundHeader As Boolean
    lr = wsKpi.Cells(wsKpi.Rows.Count, 4).End(xlUp).Row
    For r = 1 To lr
        If CStr(wsKpi.Cells(r, 4).Value) = "AI ALERTS" Then foundHeader = True: r = r + 1
        If foundHeader Then
            If CStr(wsKpi.Cells(r, 4).Value) = "AI RECOMMENDATIONS" Then Exit For
            If r > 2 And Len(Trim$(CStr(wsKpi.Cells(r, 4).Value))) > 0 Then
                If LCase$(CStr(wsKpi.Cells(r, 4).Value)) <> "severity" Then
                    alertRows.Add Array(wsKpi.Cells(r, 4).Value, _
                                        wsKpi.Cells(r, 5).Value, _
                                        wsKpi.Cells(r, 6).Value, _
                                        wsKpi.Cells(r, 7).Value)
                End If
            End If
        End If
    Next r

    ' Render up to 5 alert pills
    Dim i As Long, n As Long, rowH As Double
    n = WorksheetFunction.Min(5, alertRows.Count)
    rowH = 36
    For i = 1 To n
        Dim rec As Variant: rec = alertRows(i)
        DrawAlertCard ws, "ai_card_" & i, _
            left + 8, top + 28 + 8 + (i - 1) * (rowH + 4), w - 16, rowH, _
            CStr(rec(0)), CStr(rec(2)), CStr(rec(3))
    Next i

    If n = 0 Then
        Dim ok As Shape
        Set ok = AddShape(ws, msoShapeRectangle, "ai_card_empty", left + 16, top + 60, w - 32, 30)
        StyleShape ok, CLR_PANEL, 0
        ok.TextFrame2.TextRange.Text = ChrW(10003) & "  ALL CLEAR  |  No active AI alerts"
        With ok.TextFrame2.TextRange.Font
            .Size = 10: .Bold = msoTrue: .Name = "Segoe UI"
            .Fill.ForeColor.RGB = CLR_GOOD
        End With
        ok.TextFrame2.HorizontalAnchor = msoAnchorCenter
        ok.TextFrame2.VerticalAnchor = msoAnchorMiddle
    End If
End Sub

Private Sub DrawAlertCard(ws As Worksheet, ByVal name As String, _
        ByVal lft As Double, ByVal tp As Double, _
        ByVal w As Double, ByVal h As Double, _
        ByVal sev As String, ByVal title As String, ByVal action As String)

    Dim col As Long
    Select Case sev
        Case "Critical": col = CLR_BAD
        Case "High":     col = CLR_WARN
        Case "Medium":   col = CLR_ACCENT2
        Case Else:       col = CLR_PANEL_HI
    End Select

    Dim card As Shape
    Set card = AddShape(ws, msoShapeRoundedRectangle, name, lft, tp, w, h)
    card.Adjustments.Item(1) = 0.2
    StyleShape card, CLR_PANEL_HI, 0

    Dim sevPill As Shape
    Set sevPill = AddShape(ws, msoShapeRoundedRectangle, name & "_pill", lft + 8, tp + 8, 60, h - 16)
    sevPill.Adjustments.Item(1) = 0.4
    StyleShape sevPill, col, 0
    sevPill.TextFrame2.TextRange.Text = UCase$(sev)
    With sevPill.TextFrame2.TextRange.Font
        .Size = 8: .Bold = msoTrue: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = vbWhite
    End With
    sevPill.TextFrame2.HorizontalAnchor = msoAnchorCenter
    sevPill.TextFrame2.VerticalAnchor = msoAnchorMiddle

    Dim ttl As Shape
    Set ttl = AddShape(ws, msoShapeRectangle, name & "_t", lft + 76, tp + 4, w - 84, 18)
    StyleShape ttl, CLR_PANEL_HI, 0
    ttl.TextFrame2.TextRange.Text = title
    With ttl.TextFrame2.TextRange.Font
        .Size = 10: .Bold = msoTrue: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_TEXT
    End With
    ttl.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

    Dim act As Shape
    Set act = AddShape(ws, msoShapeRectangle, name & "_a", lft + 76, tp + 18, w - 84, h - 22)
    StyleShape act, CLR_PANEL_HI, 0
    act.TextFrame2.TextRange.Text = action
    With act.TextFrame2.TextRange.Font
        .Size = 8.5: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_MUTED
    End With
    act.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
End Sub

'==============================================================================
'                       LIVE ACTIVITY CONSOLE
' Shows latest 8 tickets like a NOC log feed
'==============================================================================
Private Sub BuildActivityConsole(ws As Worksheet)
    Dim left As Double, top As Double, w As Double, h As Double
    left = ws.Cells(G_BODY_ROW + 16, 18).Left
    top = ws.Cells(G_BODY_ROW + 16, 18).Top
    w = ws.Range("R:AB").Width
    h = 195

    Dim hdr As Shape
    Set hdr = AddShape(ws, msoShapeRectangle, "act_hdr", left, top, w, 24)
    StyleShape hdr, CLR_BG, 0
    hdr.TextFrame2.TextRange.Text = "LIVE ACTIVITY CONSOLE  " & ChrW(9656)
    With hdr.TextFrame2.TextRange.Font
        .Bold = msoTrue: .Size = 11: .Name = "Segoe UI Semibold"
        .Fill.ForeColor.RGB = CLR_ACCENT
    End With

    Dim pnl As Shape
    Set pnl = AddShape(ws, msoShapeRoundedRectangle, "act_pnl", left, top + 28, w, h)
    pnl.Adjustments.Item(1) = 0.04
    StyleShape pnl, CLR_PANEL, 0
    AddShadow pnl

    ' Pull latest 8 tickets from Clean_Data sorted by Reported_DT desc
    Dim arr As Variant
    arr = ReadCleanArrayPub
    If IsEmpty(arr) Then Exit Sub
    Dim n As Long: n = UBound(arr, 1)
    Dim idx() As Long, dts() As Double, i As Long, j As Long, t As Long, td As Double
    ReDim idx(2 To n): ReDim dts(2 To n)
    For i = 2 To n
        idx(i) = i
        If IsDate(arr(i, CD_REPORT_DT)) Then dts(i) = CDbl(CDate(arr(i, CD_REPORT_DT))) Else dts(i) = 0
    Next i
    For i = 2 To n - 1
        For j = i + 1 To n
            If dts(j) > dts(i) Then
                td = dts(i): dts(i) = dts(j): dts(j) = td
                t = idx(i): idx(i) = idx(j): idx(j) = t
            End If
        Next j
    Next i

    Dim show As Long: show = WorksheetFunction.Min(8, n - 1)
    Dim rowH As Double: rowH = 18
    For i = 1 To show
        Dim r As Long: r = idx(1 + i)
        DrawActivityRow ws, "act_r" & i, _
            left + 8, top + 28 + 6 + (i - 1) * rowH, w - 16, rowH, _
            CStr(arr(r, CD_NORM_SEVERITY)), _
            CStr(arr(r, CD_ISSUE_ID)), _
            CStr(arr(r, CD_ISSUE_TITLE)), _
            CStr(arr(r, CD_NORM_STATUS)), _
            arr(r, CD_REPORT_DT)
    Next i
End Sub

Private Sub DrawActivityRow(ws As Worksheet, ByVal name As String, _
        ByVal lft As Double, ByVal tp As Double, _
        ByVal w As Double, ByVal h As Double, _
        ByVal sev As String, ByVal id As String, ByVal title As String, _
        ByVal stat As String, ByVal ts As Variant)

    Dim col As Long
    Select Case sev
        Case "Critical": col = CLR_BAD
        Case "High":     col = CLR_WARN
        Case "Medium":   col = CLR_ACCENT2
        Case "Low":      col = CLR_GOOD
        Case Else:       col = CLR_MUTED
    End Select

    ' Severity dot
    Dim dot As Shape
    Set dot = AddShape(ws, msoShapeOval, name & "_d", lft + 4, tp + 4, 10, 10)
    StyleShape dot, col, 0

    ' Text strip
    Dim t As Shape
    Set t = AddShape(ws, msoShapeRectangle, name & "_t", lft + 18, tp, w - 18, h)
    StyleShape t, CLR_PANEL, 0
    Dim tsTxt As String
    If IsDate(ts) Then tsTxt = Format(CDate(ts), "dd-mmm hh:mm") Else tsTxt = ""
    t.TextFrame2.TextRange.Text = "#" & id & "   " & TrimTo(title, 38) & "   |  " & stat & "  |  " & tsTxt
    With t.TextFrame2.TextRange.Font
        .Size = 9: .Name = "Consolas"
        .Fill.ForeColor.RGB = CLR_TEXT
    End With
    t.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    t.TextFrame2.VerticalAnchor = msoAnchorMiddle
End Sub

'==============================================================================
'                          FOOTER BAR
'==============================================================================
Private Sub BuildFooterBar(ws As Worksheet)
    Dim btnLeft As Double, btnTop As Double, btnW As Double, btnH As Double
    Dim row As Long: row = 35
    btnLeft = ws.Cells(row, 2).Left
    btnTop = ws.Cells(row, 2).Top
    btnW = 130: btnH = 30

    ' Refresh Button (left-most)
    AddCommandButton ws, "btn_refresh", btnLeft, btnTop, btnW, btnH, _
        ChrW(8635) & "  REFRESH", "RefreshAll", CLR_ACCENT, CLR_BG

    AddCommandButton ws, "btn_export", btnLeft + 140, btnTop, btnW, btnH, _
        ChrW(8595) & "  EXPORT PDF", "ExportDashboardPdf", CLR_ACCENT2, CLR_BG

    AddCommandButton ws, "btn_email", btnLeft + 280, btnTop, btnW, btnH, _
        ChrW(9993) & "  EMAIL MIS", "EmailMisReport", CLR_PURPLE, vbWhite

    AddCommandButton ws, "btn_theme", btnLeft + 420, btnTop, btnW, btnH, _
        ChrW(9789) & "  THEME", "ToggleTheme", CLR_PANEL_HI, CLR_TEXT

    AddCommandButton ws, "btn_search", btnLeft + 560, btnTop, btnW, btnH, _
        ChrW(128269) & "  SEARCH", "SmartSearch", CLR_PANEL_HI, CLR_TEXT

    ' Footer text
    Dim foot As Shape
    Set foot = AddShape(ws, msoShapeRectangle, "foot_txt", btnLeft, btnTop + 40, 800, 20)
    StyleShape foot, CLR_BG, 0
    foot.TextFrame2.TextRange.Text = APP_NAME & "  v" & APP_VERSION & "   |   Built " & Format(Now, "dd-mmm-yyyy hh:mm")
    With foot.TextFrame2.TextRange.Font
        .Size = 8: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_MUTED
    End With
End Sub

'==============================================================================
'                          SHAPE HELPERS (shared across modules)
'==============================================================================
Public Function AddShape(ws As Worksheet, ByVal sType As MsoAutoShapeType, ByVal name As String, _
        ByVal lft As Double, ByVal tp As Double, ByVal w As Double, ByVal h As Double) As Shape
    Dim sh As Shape
    On Error Resume Next
    ws.Shapes(name).Delete
    On Error GoTo 0
    Set sh = ws.Shapes.AddShape(sType, lft, tp, w, h)
    sh.Name = name
    Set AddShape = sh
End Function

Public Sub StyleShape(sh As Shape, ByVal fillColor As Long, ByVal lineColor As Long)
    sh.Fill.Visible = msoTrue
    sh.Fill.ForeColor.RGB = fillColor
    sh.Fill.Solid
    If lineColor = 0 Then
        sh.Line.Visible = msoFalse
    Else
        sh.Line.ForeColor.RGB = lineColor
        sh.Line.Weight = 1
    End If
    sh.TextFrame2.MarginLeft = 6
    sh.TextFrame2.MarginRight = 6
    sh.TextFrame2.MarginTop = 2
    sh.TextFrame2.MarginBottom = 2
End Sub

Public Sub AddShadow(sh As Shape)
    With sh.Shadow
        .Visible = msoTrue
        .Type = msoShadow21
        .Blur = 8
        .OffsetX = 0
        .OffsetY = 2
        .Transparency = 0.6
    End With
End Sub

Public Sub DeleteShapesByPrefix(ws As Worksheet, ByVal prefix As String)
    Dim sh As Shape, i As Long
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If LCase$(Left$(sh.Name, Len(prefix))) = LCase$(prefix) Then sh.Delete
    Next i
End Sub

Public Function TrimTo(ByVal s As String, ByVal n As Long) As String
    If Len(s) <= n Then TrimTo = s Else TrimTo = Left$(s, n - 1) & ChrW(8230)
End Function


'==============================================================================
' === SECTION: M05_Charts ===
'==============================================================================
Public Sub RefreshAllCharts()
    ' Rebuild only the embedded chart objects on Dashboard_Main and module sheets
    ' Layout itself is preserved.  This is a thin wrapper used by RefreshAll.
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Visible = xlSheetVisible Then
            On Error Resume Next
            ws.ChartObjects.Delete
            On Error GoTo 0
        End If
    Next ws
    BuildDashboardMain
    RefreshAllModules
End Sub

'==============================================================================
'                          DASHBOARD CORE CHARTS
'==============================================================================
Public Sub BuildTrendChart(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double, ByVal title As String)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Daily Volume", r1, r2
    If r1 = 0 Then Exit Sub

    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_trend", lft, tp, w, h)
    With co.Chart
        .ChartType = xlLineMarkers
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        Dim dataRange As Range
        Set dataRange = dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 4))
        .SetSourceData Source:=dataRange, PlotBy:=xlColumns
        StyleChartDark .Parent.Chart, title
        ' Series styling
        On Error Resume Next
        Dim s1 As Series, s2 As Series, s3 As Series
        Set s1 = .SeriesCollection(1)  ' Count
        Set s2 = .SeriesCollection(2)  ' MA(3)
        Set s3 = .SeriesCollection(3)  ' Forecast
        StyleSeriesLine s1, CLR_ACCENT, 2.5, True
        StyleSeriesLine s2, CLR_ACCENT2, 1.75, False
        StyleSeriesLine s3, CLR_PURPLE, 2#, True
        s3.Format.Line.DashStyle = msoLineDash
        On Error GoTo 0
    End With
End Sub

Public Sub BuildSeverityDonut(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double, ByVal title As String)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Severity", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_sev", lft, tp, w, h)
    With co.Chart
        .ChartType = xlDoughnut
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        .SetSourceData Source:=dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 2))
        StyleChartDark .Parent.Chart, title
        ' Color slices
        Dim pts As Points, i As Long
        Set pts = .SeriesCollection(1).Points
        For i = 1 To pts.count
            Dim k As String
            k = CStr(dm.Cells(r1 + i, 1).Value)
            pts(i).Format.Fill.Visible = msoTrue
            pts(i).Format.Fill.ForeColor.RGB = SeverityColor(k)
            pts(i).Format.Line.Visible = msoFalse
        Next i
        ' Big hole
        .SeriesCollection(1).DoughnutHoleSize = 70
        ' Data labels
        .SeriesCollection(1).HasDataLabels = True
        With .SeriesCollection(1).DataLabels
            .ShowValue = True
            .ShowCategoryName = True
            .Separator = "  "
            .Font.Color = CLR_TEXT
            .Font.Size = 9
            .Font.Name = "Segoe UI"
        End With
    End With
End Sub

Public Sub BuildAgingFunnel(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double, ByVal title As String)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Aging Bucket", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_aging", lft, tp, w, h)
    With co.Chart
        .ChartType = xlBarClustered
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        .SetSourceData Source:=dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 2))
        StyleChartDark .Parent.Chart, title
        Dim s As Series: Set s = .SeriesCollection(1)
        s.Format.Fill.ForeColor.RGB = CLR_ACCENT
        s.Format.Line.Visible = msoFalse
        ' Color gradient by bucket position
        Dim pts As Points, i As Long
        Set pts = s.Points
        For i = 1 To pts.count
            pts(i).Format.Fill.Visible = msoTrue
            pts(i).Format.Fill.ForeColor.RGB = AgingColor(CStr(dm.Cells(r1 + i, 1).Value))
        Next i
        s.HasDataLabels = True
        With s.DataLabels
            .Font.Color = CLR_TEXT: .Font.Size = 9: .Font.Name = "Segoe UI"
            .Position = xlLabelPositionInsideEnd
        End With
    End With
End Sub

Public Sub BuildCategoryBars(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double, ByVal title As String)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Category", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_cat", lft, tp, w, h)
    With co.Chart
        .ChartType = xlColumnClustered
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        Dim r2c As Long: r2c = WorksheetFunction.Min(r2, r1 + 8)  ' cap to top 8
        .SetSourceData Source:=dm.Range(dm.Cells(r1, 1), dm.Cells(r2c, 2))
        StyleChartDark .Parent.Chart, title
        Dim s As Series: Set s = .SeriesCollection(1)
        s.Format.Fill.ForeColor.RGB = CLR_ACCENT2
        s.Format.Line.Visible = msoFalse
        s.HasDataLabels = True
        With s.DataLabels
            .Font.Color = CLR_TEXT: .Font.Size = 9: .Font.Name = "Segoe UI"
            .Position = xlLabelPositionOutsideEnd
        End With
        On Error Resume Next
        .ChartArea.Format.TextFrame2.TextRange.Font.Size = 9
        On Error GoTo 0
    End With
End Sub

'==============================================================================
'                          MODULE-LEVEL CHARTS
'==============================================================================
Public Sub BuildHourHeatChart(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    ' Render hour-of-day distribution as a stylized column chart
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Hour Of Day", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_hour_" & RandSuffix, lft, tp, w, h)
    With co.Chart
        .ChartType = xlColumnClustered
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        .SetSourceData dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 2))
        StyleChartDark .Parent.Chart, "Peak Hour Distribution"
        Dim s As Series: Set s = .SeriesCollection(1)
        s.Format.Fill.ForeColor.RGB = CLR_ACCENT
        s.Format.Line.Visible = msoFalse
        s.GapWidth = 30
    End With
End Sub

Public Sub BuildDowChart(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Day Of Week", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_dow_" & RandSuffix, lft, tp, w, h)
    With co.Chart
        .ChartType = xlColumnClustered
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        .SetSourceData dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 2))
        StyleChartDark .Parent.Chart, "Day-of-Week Volume"
        Dim s As Series: Set s = .SeriesCollection(1)
        s.Format.Fill.ForeColor.RGB = CLR_ACCENT2
        s.Format.Line.Visible = msoFalse
    End With
End Sub

Public Sub BuildSlaGauge(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    ' Semi-donut gauge using two synthetic data points: compliance + remainder
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_gauge", lft, tp, w, h)
    Dim hidden As Worksheet: Set hidden = ThisWorkbook.Worksheets(SHT_HIDDEN)
    Dim slaVal As Variant
    slaVal = GetKpi("SLA_Compliance")
    Dim val As Double: val = NzNum(slaVal)
    If val > 1 Then val = val / 100
    hidden.Range("AA1:AA4").Value = Application.Transpose(Array("Met", val, "Gap", 1 - val))
    hidden.Range("AB1:AB4").Value = Application.Transpose(Array(val, 1 - val, 1, 1))   ' bottom half hidden

    With co.Chart
        .ChartType = xlDoughnut
        .SetSourceData hidden.Range("AA1:AA4")
        StyleChartDark .Parent.Chart, "SLA Compliance " & Format(val, "0.0%")
        Dim ser As Series: Set ser = .SeriesCollection(1)
        ser.DoughnutHoleSize = 65
        ' Color points: 1=met (green), 2=gap (red), 3,4 hidden
        ser.Points(1).Format.Fill.ForeColor.RGB = IIf(val >= 0.85, CLR_GOOD, IIf(val >= 0.7, CLR_WARN, CLR_BAD))
        ser.Points(2).Format.Fill.ForeColor.RGB = CLR_PANEL_HI
        ser.Points(3).Format.Fill.Visible = msoFalse
        ser.Points(4).Format.Fill.Visible = msoFalse
        ser.Points(1).Format.Line.Visible = msoFalse
        ser.Points(2).Format.Line.Visible = msoFalse
        ' Rotate so gauge is bottom-half-style
        .SetElement msoElementLegendNone
        ser.HasDataLabels = False
        .ChartArea.Format.TextFrame2.TextRange.Font.Color.RGB = CLR_TEXT
    End With
End Sub

Public Sub BuildSlaSeverityChart(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: SLA by Severity", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_slasev", lft, tp, w, h)
    With co.Chart
        .ChartType = xlColumnStacked
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        ' Severity, Met, At Risk, Breached  -> cols 1,3,4,5
        Dim src As Range
        Set src = Application.Union( _
            dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 1)), _
            dm.Range(dm.Cells(r1, 3), dm.Cells(r2, 5)))
        .SetSourceData src
        StyleChartDark .Parent.Chart, "SLA Status by Severity"
        On Error Resume Next
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = CLR_GOOD
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = CLR_WARN
        .SeriesCollection(3).Format.Fill.ForeColor.RGB = CLR_BAD
        Dim i As Long
        For i = 1 To .SeriesCollection.count
            .SeriesCollection(i).Format.Line.Visible = msoFalse
        Next i
        On Error GoTo 0
    End With
End Sub

Public Sub BuildResolverChart(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Resolver MTTR", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_resolver", lft, tp, w, h)
    With co.Chart
        .ChartType = xlBarClustered
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        Dim r2c As Long: r2c = WorksheetFunction.Min(r2, r1 + 10)
        ' Resolver + Tickets
        .SetSourceData dm.Range(dm.Cells(r1, 1), dm.Cells(r2c, 2))
        StyleChartDark .Parent.Chart, "Top Resolvers - Ticket Volume"
        Dim s As Series: Set s = .SeriesCollection(1)
        s.Format.Fill.ForeColor.RGB = CLR_ACCENT
        s.Format.Line.Visible = msoFalse
        s.HasDataLabels = True
        s.DataLabels.Font.Color = CLR_TEXT
        s.DataLabels.Font.Size = 9
    End With
End Sub

Public Sub BuildResolverEfficiencyChart(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Resolver MTTR", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_reseff", lft, tp, w, h)
    With co.Chart
        .ChartType = xlXYScatter
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        Dim r2c As Long: r2c = WorksheetFunction.Min(r2, r1 + 10)
        ' X = avg hours (col 3), Y = SLA met % (col 4), bubble size = tickets (col 2)
        .SeriesCollection.NewSeries
        With .SeriesCollection(1)
            .XValues = dm.Range(dm.Cells(r1 + 1, 3), dm.Cells(r2c, 3))
            .Values = dm.Range(dm.Cells(r1 + 1, 4), dm.Cells(r2c, 4))
            .name = "Resolver Efficiency"
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 12
            .MarkerForegroundColor = CLR_ACCENT
            .MarkerBackgroundColor = CLR_ACCENT
            .Format.Line.Visible = msoFalse
        End With
        StyleChartDark .Parent.Chart, "Resolver Speed vs SLA Met %"
        .Axes(xlCategory).HasTitle = True
        .Axes(xlCategory).AxisTitle.Text = "Avg Resolution (hours)"
        .Axes(xlValue).HasTitle = True
        .Axes(xlValue).AxisTitle.Text = "SLA Met %"
        .Axes(xlCategory).AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_MUTED
        .Axes(xlValue).AxisTitle.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_MUTED
    End With
End Sub

Public Sub BuildLocationRiskChart(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Risk by Location", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_riskloc", lft, tp, w, h)
    With co.Chart
        .ChartType = xlBarClustered
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        Dim r2c As Long: r2c = WorksheetFunction.Min(r2, r1 + 10)
        ' Location + Avg Risk
        Dim src As Range
        Set src = Application.Union( _
            dm.Range(dm.Cells(r1, 1), dm.Cells(r2c, 1)), _
            dm.Range(dm.Cells(r1, 3), dm.Cells(r2c, 3)))
        .SetSourceData src
        StyleChartDark .Parent.Chart, "Top High-Risk Locations"
        Dim s As Series: Set s = .SeriesCollection(1)
        s.Format.Fill.ForeColor.RGB = CLR_BAD
        s.Format.Line.Visible = msoFalse
        s.HasDataLabels = True
        s.DataLabels.Font.Color = CLR_TEXT
        s.DataLabels.Font.Size = 9
        s.DataLabels.NumberFormat = "0.0"
    End With
End Sub

Public Sub BuildIssueTypeChart(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: IssueType", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_itype", lft, tp, w, h)
    With co.Chart
        .ChartType = xlPie
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        .SetSourceData dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 2))
        StyleChartDark .Parent.Chart, "Issue Type Mix"
        Dim pts As Points, i As Long, palette As Variant
        palette = Array(CLR_ACCENT, CLR_ACCENT2, CLR_PURPLE, CLR_PINK, CLR_GOOD, CLR_WARN, CLR_BAD)
        Set pts = .SeriesCollection(1).Points
        For i = 1 To pts.count
            pts(i).Format.Fill.ForeColor.RGB = CLng(palette((i - 1) Mod 7))
            pts(i).Format.Line.Visible = msoFalse
        Next i
        .SeriesCollection(1).HasDataLabels = True
        With .SeriesCollection(1).DataLabels
            .ShowPercentage = True
            .ShowCategoryName = True
            .Separator = "  "
            .Font.Color = CLR_TEXT: .Font.Size = 9
        End With
    End With
End Sub

Public Sub BuildLocationVolumeChart(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Top 10 Reporter Locations", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_loc", lft, tp, w, h)
    With co.Chart
        .ChartType = xlBarClustered
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        .SetSourceData dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 2))
        StyleChartDark .Parent.Chart, "Top 10 Locations by Tickets"
        Dim s As Series: Set s = .SeriesCollection(1)
        s.Format.Fill.ForeColor.RGB = CLR_PURPLE
        s.Format.Line.Visible = msoFalse
        s.HasDataLabels = True
        s.DataLabels.Font.Color = CLR_TEXT
        s.DataLabels.Font.Size = 9
    End With
End Sub

Public Sub BuildForecastChart(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Daily Volume", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_fc", lft, tp, w, h)
    With co.Chart
        .ChartType = xlAreaStacked
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        ' Date + Count + Forecast
        Dim src As Range
        Set src = Application.Union( _
            dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 2)), _
            dm.Range(dm.Cells(r1, 4), dm.Cells(r2, 4)))
        .SetSourceData src
        StyleChartDark .Parent.Chart, "Volume + 3-Day Forecast"
        On Error Resume Next
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = CLR_ACCENT
        .SeriesCollection(1).Format.Line.Visible = msoFalse
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = CLR_PURPLE
        .SeriesCollection(2).Format.Line.Visible = msoFalse
        On Error GoTo 0
    End With
End Sub

Public Sub BuildOpsRadar(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    ' 6-axis radar showing operational health dimensions
    Dim hidden As Worksheet: Set hidden = ThisWorkbook.Worksheets(SHT_HIDDEN)
    hidden.Range("AD1:AE7").ClearContents
    hidden.Range("AD1:AE1").Value = Array("Dimension", "Score")
    hidden.Range("AD2").Value = "SLA":             hidden.Range("AE2").Value = NzNum(GetKpi("SLA_Compliance")) * 100
    hidden.Range("AD3").Value = "Stability":       hidden.Range("AE3").Value = NzNum(GetKpi("Stability_Score"))
    hidden.Range("AD4").Value = "Closure":         hidden.Range("AE4").Value = NzNum(GetKpi("ClosureRate")) * 100
    hidden.Range("AD5").Value = "Efficiency":      hidden.Range("AE5").Value = NzNum(GetKpi("Agent_Efficiency"))
    hidden.Range("AD6").Value = "CSAT":            hidden.Range("AE6").Value = NzNum(GetKpi("Customer_Satisfaction"))
    hidden.Range("AD7").Value = "Health":          hidden.Range("AE7").Value = NzNum(GetKpi("OpsHealth_Score"))

    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_radar", lft, tp, w, h)
    With co.Chart
        .ChartType = xlRadarFilled
        .SetSourceData hidden.Range("AD1:AE7")
        StyleChartDark .Parent.Chart, "Operational Health Radar"
        Dim s As Series: Set s = .SeriesCollection(1)
        s.Format.Fill.ForeColor.RGB = CLR_ACCENT
        s.Format.Fill.Transparency = 0.5
        s.Format.Line.ForeColor.RGB = CLR_ACCENT
        s.Format.Line.Weight = 2
        s.HasDataLabels = True
        s.DataLabels.Font.Color = CLR_TEXT
        s.DataLabels.Font.Size = 9
    End With
End Sub

Public Sub BuildSeverityStatusHeat(ws As Worksheet, lft As Double, tp As Double, _
        w As Double, h As Double)
    ' Render the severity x status cross-tab as a stacked bar (poor man's heatmap)
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Severity x Status (Heatmap)", r1, r2
    If r1 = 0 Then Exit Sub
    Dim co As ChartObject
    Set co = AddChartObject(ws, "ch_sevstat", lft, tp, w, h)
    With co.Chart
        .ChartType = xlBarStacked100
        Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
        .SetSourceData dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 3))
        StyleChartDark .Parent.Chart, "Severity x Status Distribution"
        On Error Resume Next
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = CLR_WARN     ' Open
        .SeriesCollection(2).Format.Fill.ForeColor.RGB = CLR_GOOD     ' Closed
        Dim i As Long
        For i = 1 To .SeriesCollection.count
            .SeriesCollection(i).Format.Line.Visible = msoFalse
        Next i
        On Error GoTo 0
    End With
End Sub

Public Sub BuildHourDowHeatmap(ws As Worksheet, ByVal anchorRow As Long, ByVal anchorCol As Long)
    ' Build a 25x8 cell grid with conditional formatting -> NOC heatmap
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Hour x DOW (Heatmap)", r1, r2
    If r1 = 0 Then Exit Sub
    Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
    Dim src As Range
    Set src = dm.Range(dm.Cells(r1, 1), dm.Cells(r2, 8))
    Dim arr As Variant: arr = src.Value

    Dim rows As Long, cols As Long
    rows = UBound(arr, 1)
    cols = UBound(arr, 2)

    Dim dst As Range
    Set dst = ws.Range(ws.Cells(anchorRow, anchorCol), ws.Cells(anchorRow + rows - 1, anchorCol + cols - 1))
    dst.Value = arr
    With dst
        .Font.Color = CLR_TEXT
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = CLR_BG
    End With
    ' Header row + first column styling
    dst.Rows(1).Font.Bold = True
    dst.Rows(1).Interior.Color = CLR_PANEL
    dst.Columns(1).Font.Bold = True
    dst.Columns(1).Interior.Color = CLR_PANEL

    ' Heat color scale on the inner numeric grid
    Dim heat As Range
    Set heat = ws.Range(ws.Cells(anchorRow + 1, anchorCol + 1), ws.Cells(anchorRow + rows - 1, anchorCol + cols - 1))
    heat.FormatConditions.Delete
    Dim cs As ColorScale
    Set cs = heat.FormatConditions.AddColorScale(ColorScaleType:=3)
    cs.ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    cs.ColorScaleCriteria(1).FormatColor.Color = CLR_PANEL
    cs.ColorScaleCriteria(2).Type = xlConditionValuePercentile
    cs.ColorScaleCriteria(2).Value = 50
    cs.ColorScaleCriteria(2).FormatColor.Color = CLR_ACCENT2
    cs.ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    cs.ColorScaleCriteria(3).FormatColor.Color = CLR_ACCENT
End Sub

'==============================================================================
'                          STYLING
'==============================================================================
Public Sub StyleChartDark(ch As Chart, ByVal title As String)
    ch.HasTitle = True
    With ch.ChartTitle
        .Text = title
        With .Format.TextFrame2.TextRange.Font
            .Bold = msoTrue
            .Size = 11
            .Name = "Segoe UI Semibold"
            .Fill.ForeColor.RGB = CLR_ACCENT
        End With
        .Left = 8
        .Top = 4
    End With

    ' Plot/chart areas
    With ch.ChartArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = CLR_PANEL
        .Solid
    End With
    ch.ChartArea.Format.Line.Visible = msoFalse

    With ch.PlotArea.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = CLR_PANEL
        .Solid
    End With
    ch.PlotArea.Format.Line.Visible = msoFalse

    ' Axes
    On Error Resume Next
    Dim ax As Axis
    For Each ax In ch.Axes
        ax.Format.Line.ForeColor.RGB = CLR_BORDER
        ax.MajorGridlines.Format.Line.ForeColor.RGB = CLR_BORDER
        ax.MajorGridlines.Format.Line.Transparency = 0.6
        ax.TickLabels.Font.Color = CLR_MUTED
        ax.TickLabels.Font.Size = 8
        ax.TickLabels.Font.Name = "Segoe UI"
    Next ax
    On Error GoTo 0

    ' Legend
    On Error Resume Next
    ch.HasLegend = True
    ch.Legend.Format.Fill.Visible = msoFalse
    ch.Legend.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = CLR_MUTED
    ch.Legend.Font.Size = 8
    ch.Legend.Position = xlLegendPositionBottom
    On Error GoTo 0
End Sub

Public Sub StyleSeriesLine(s As Series, ByVal col As Long, ByVal weight As Double, ByVal showMarker As Boolean)
    s.Format.Line.ForeColor.RGB = col
    s.Format.Line.Weight = weight
    If showMarker Then
        On Error Resume Next
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 6
        s.MarkerForegroundColor = col
        s.MarkerBackgroundColor = col
        On Error GoTo 0
    End If
End Sub

'==============================================================================
'                          HELPERS
'==============================================================================
Public Function AddChartObject(ws As Worksheet, ByVal name As String, _
        ByVal lft As Double, ByVal tp As Double, _
        ByVal w As Double, ByVal h As Double) As ChartObject
    Dim co As ChartObject
    On Error Resume Next
    ws.ChartObjects(name).Delete
    On Error GoTo 0
    Set co = ws.ChartObjects.Add(lft, tp, w, h)
    co.name = name
    Set AddChartObject = co
End Function

Public Sub LocateBlock(ByVal hdrText As String, ByRef firstHeaderRow As Long, ByRef lastDataRow As Long)
    ' Finds "AGG: ..." header on Data_Model and returns the column-header row
    ' (one row below the title) and the last data row (until next title or blank).
    firstHeaderRow = 0: lastDataRow = 0
    Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)
    Dim r As Long, lr As Long
    lr = dm.Cells(dm.Rows.count, 1).End(xlUp).Row
    For r = 1 To lr
        If StrComp(CStr(dm.Cells(r, 1).Value), hdrText, vbTextCompare) = 0 Then
            firstHeaderRow = r + 1
            Exit For
        End If
    Next r
    If firstHeaderRow = 0 Then Exit Sub
    Dim r2 As Long: r2 = firstHeaderRow + 1
    Do While r2 <= lr
        Dim v As String: v = CStr(dm.Cells(r2, 1).Value)
        If Len(v) = 0 Then Exit Do
        If Left$(v, 4) = "AGG:" Then Exit Do
        r2 = r2 + 1
    Loop
    lastDataRow = r2 - 1
End Sub

Public Function SeverityColor(ByVal sev As String) As Long
    Select Case sev
        Case "Critical": SeverityColor = CLR_BAD
        Case "High":     SeverityColor = CLR_WARN
        Case "Medium":   SeverityColor = CLR_ACCENT2
        Case "Low":      SeverityColor = CLR_GOOD
        Case Else:       SeverityColor = CLR_MUTED
    End Select
End Function

Public Function AgingColor(ByVal bucket As String) As Long
    Select Case bucket
        Case "00-04h": AgingColor = CLR_GOOD
        Case "04-24h": AgingColor = CLR_ACCENT2
        Case "1-3d":   AgingColor = CLR_WARN
        Case "3-7d":   AgingColor = RGB(255, 127, 0)   ' orange
        Case ">7d":    AgingColor = CLR_BAD
        Case Else:     AgingColor = CLR_MUTED
    End Select
End Function


Private Function RandSuffix() As String
    RandSuffix = Format(Now, "hhnnss") & Int(Rnd * 1000)
End Function
