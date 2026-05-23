Attribute VB_Name = "M04_DashboardUI"
'==============================================================================
' MODULE      : M04_DashboardUI
' DESCRIPTION : Builds the executive Dashboard_Main page with premium dark NOC
'               aesthetic: header bar, side navigation, KPI cards, AI feed,
'               activity console and chart hosts. Uses Excel shapes for visual
'               primitives (cards, panels, neon accents) and embedded charts.
'==============================================================================
Option Explicit
Option Compare Text

' Dashboard layout grid (cells)
Private Const G_TITLE_ROW   As Long = 1   ' header
Private Const G_KPI_ROW     As Long = 4   ' kpi cards row anchor
Private Const G_BODY_ROW    As Long = 11  ' body content begins
Private Const G_NAV_COL     As Long = 1
Private Const G_BODY_COL    As Long = 2

'==============================================================================
'                          PUBLIC ENTRY POINTS
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
    M07_Interaction.BuildNavigationRail ws, SHT_DASH

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
    Set ws = ThisWorkbook.Worksheets(SHT_DASH)
    ' Repaint KPI cards by clearing+rebuilding shapes prefixed kpi_*
    DeleteShapesByPrefix ws, "kpi_"
    DeleteShapesByPrefix ws, "ai_"
    DeleteShapesByPrefix ws, "act_"
    BuildPrimaryKpiCards ws
    BuildSecondaryKpiStrip ws
    BuildAiAlertPanel ws
    BuildActivityConsole ws
    M05_Charts.RefreshAllCharts
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
    val.TextFrame2.TextRange.Text = M03_KpiEngine.FmtKpi(kpiKey, numFmt)
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
    val.TextFrame2.TextRange.Text = M03_KpiEngine.FmtKpi(kpiKey, numFmt)
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

    M05_Charts.BuildTrendChart ws, chartLeft, chartTop, chartW, chartH, "Daily Ticket Volume + Forecast"
    M05_Charts.BuildSeverityDonut ws, chartLeft + chartW + gap, chartTop, chartW, chartH, "Severity Distribution"
    M05_Charts.BuildAgingFunnel ws, chartLeft, chartTop + chartH + gap, chartW, chartH, "Ticket Aging Buckets"
    M05_Charts.BuildCategoryBars ws, chartLeft + chartW + gap, chartTop + chartH + gap, chartW, chartH, "Category Heat (Auto-Clustered)"
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
    arr = M03_KpiEngine.ReadCleanArrayPub
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
    M07_Interaction.AddCommandButton ws, "btn_refresh", btnLeft, btnTop, btnW, btnH, _
        ChrW(8635) & "  REFRESH", "M01_Builder.RefreshAll", CLR_ACCENT, CLR_BG

    M07_Interaction.AddCommandButton ws, "btn_export", btnLeft + 140, btnTop, btnW, btnH, _
        ChrW(8595) & "  EXPORT PDF", "M07_Interaction.ExportDashboardPdf", CLR_ACCENT2, CLR_BG

    M07_Interaction.AddCommandButton ws, "btn_email", btnLeft + 280, btnTop, btnW, btnH, _
        ChrW(9993) & "  EMAIL MIS", "M07_Interaction.EmailMisReport", CLR_PURPLE, vbWhite

    M07_Interaction.AddCommandButton ws, "btn_theme", btnLeft + 420, btnTop, btnW, btnH, _
        ChrW(9789) & "  THEME", "M07_Interaction.ToggleTheme", CLR_PANEL_HI, CLR_TEXT

    M07_Interaction.AddCommandButton ws, "btn_search", btnLeft + 560, btnTop, btnW, btnH, _
        ChrW(128269) & "  SEARCH", "M07_Interaction.SmartSearch", CLR_PANEL_HI, CLR_TEXT

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
