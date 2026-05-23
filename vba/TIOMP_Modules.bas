Attribute VB_Name = "TIOMP_Modules"
Option Explicit
Option Compare Text

' === TIOMP_Modules - Part of TIOMP Dashboard ===
' Imports: M06_Modules, M07_Interaction


'==============================================================================
' === SECTION: M06_Modules ===
'==============================================================================

'==============================================================================
' MODULE      : M06_Modules
' DESCRIPTION : Builds the 8 specialized analytics module sheets that the user
'               navigates between via the side rail.  Each module follows a
'               common "Hero Header > KPI Strip > Charts > Drill Table" layout
'               so the experience feels like Power BI tabs.
'==============================================================================

'==============================================================================
'                          PUBLIC ORCHESTRATORS
'==============================================================================
Public Sub RefreshAllModules()
    BuildExecutiveView
    BuildSlaIntelligence
    BuildIncidentAnalytics
    BuildAgentAnalytics
    BuildApplicationAnalytics
    BuildRcaAnalytics
    BuildForecastAnalytics
    BuildRiskAnalytics
End Sub

'==============================================================================
'                          1. EXECUTIVE VIEW
'==============================================================================
Public Sub BuildExecutiveView()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_EXEC)
    ResetSheetForDashboard ws
    PaintCanvas ws
    BuildHeaderBar ws, "EXECUTIVE VIEW", _
        "Board-level operational scorecard with health radar"
    BuildNavigationRail ws, SHT_EXEC

    ' Hero KPIs (executive-relevant)
    Dim cards As Variant
    cards = Array( _
        Array("Ops Health", "OpsHealth_Score", "0", CLR_GOOD, ChrW(9829)), _
        Array("Stability", "Stability_Score", "0", CLR_ACCENT, ChrW(9776)), _
        Array("SLA", "SLA_Compliance", "0.0%", CLR_GOOD, ChrW(10003)), _
        Array("Risk Score", "Risk_Score", "0", CLR_BAD, ChrW(9889)), _
        Array("CSAT Index", "Customer_Satisfaction", "0", CLR_PURPLE, ChrW(9786)), _
        Array("Trend WoW %", "Trend_Velocity", "0.0", CLR_ACCENT2, ChrW(8599)))
    DrawCardRow ws, "exe", cards, 4, 6

    ' Charts
    BuildOpsRadar ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 380, 260
    BuildTrendChart ws, _
        ws.Cells(11, 12).Left, ws.Cells(11, 12).Top, 480, 260, "Daily Volume + 3-Day Forecast"
    BuildSlaGauge ws, _
        ws.Cells(11, 24).Left, ws.Cells(11, 24).Top, 240, 260
    BuildSeverityDonut ws, _
        ws.Cells(26, 2).Left, ws.Cells(26, 2).Top, 380, 240, "Severity Mix"
    BuildIssueTypeChart ws, _
        ws.Cells(26, 12).Left, ws.Cells(26, 12).Top, 380, 240
    BuildLocationVolumeChart ws, _
        ws.Cells(26, 21).Left, ws.Cells(26, 21).Top, 360, 240
End Sub

'==============================================================================
'                          2. SLA INTELLIGENCE
'==============================================================================
Public Sub BuildSlaIntelligence()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_SLA)
    ResetSheetForDashboard ws
    PaintCanvas ws
    BuildHeaderBar ws, "SLA INTELLIGENCE CENTER", _
        "Compliance, breach prediction and recovery insights"
    BuildNavigationRail ws, SHT_SLA

    Dim cards As Variant
    cards = Array( _
        Array("SLA %", "SLA_Compliance", "0.0%", CLR_GOOD, ChrW(10003)), _
        Array("Breached", "BreachedTickets", "0", CLR_BAD, ChrW(9888)), _
        Array("At Risk", "AtRiskTickets", "0", CLR_WARN, ChrW(9203)), _
        Array("MTTR (h)", "MTTR_Hours", "0.0", CLR_ACCENT2, ChrW(8987)), _
        Array("Ack (h)", "Ack_Hours", "0.0", CLR_ACCENT, ChrW(9991)), _
        Array("Crit Open", "CriticalOpen", "0", CLR_BAD, ChrW(9760)))
    DrawCardRow ws, "sla", cards, 4, 6

    BuildSlaGauge ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 280, 260
    BuildSlaSeverityChart ws, _
        ws.Cells(11, 9).Left, ws.Cells(11, 9).Top, 480, 260
    BuildAgingFunnel ws, _
        ws.Cells(11, 21).Left, ws.Cells(11, 21).Top, 380, 260, "SLA Aging Buckets"

    ' Predicted breaches drill-down
    BuildPredictedBreachesTable ws, 26, 2
End Sub

Private Sub BuildPredictedBreachesTable(ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)
    PaintSectionTitle ws, startRow, startCol, 18, "PREDICTED SLA BREACHES (open tickets)"

    Dim wsKpi As Worksheet: Set wsKpi = ThisWorkbook.Worksheets(SHT_KPI)
    Dim r As Long, lr As Long, found As Boolean
    lr = wsKpi.Cells(wsKpi.Rows.count, 4).End(xlUp).Row

    ' Headers
    Dim hdrRow As Long: hdrRow = startRow + 1
    Dim headers As Variant: headers = Array("Issue ID", "Title", "Age (h)", "SLA (h)", "Risk Pill")
    Dim i As Long
    For i = 0 To UBound(headers)
        ws.Cells(hdrRow, startCol + i).Value = headers(i)
    Next i
    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(hdrRow, startCol + 4))
        .Font.Color = CLR_ACCENT
        .Font.Bold = True
        .Interior.Color = CLR_PANEL
    End With

    ' Locate "PREDICTED SLA BREACHES" header in KPI sheet
    Dim startKpi As Long
    For r = 1 To lr
        If CStr(wsKpi.Cells(r, 4).Value) = "PREDICTED SLA BREACHES" Then startKpi = r + 2: Exit For
    Next r
    If startKpi = 0 Then Exit Sub

    Dim outRow As Long: outRow = hdrRow + 1
    For r = startKpi To lr
        Dim v As String: v = CStr(wsKpi.Cells(r, 4).Value)
        If Len(v) = 0 Then Exit For
        If UCase$(Left$(v, 4)) = "HIGH" And InStr(v, "RISK") > 0 Then Exit For
        If UCase$(v) = "HIGH-RISK LOCATIONS" Or UCase$(v) = "AI RECOMMENDATIONS" Then Exit For
        ws.Cells(outRow, startCol).Value = wsKpi.Cells(r, 4).Value
        ws.Cells(outRow, startCol + 1).Value = wsKpi.Cells(r, 5).Value
        ws.Cells(outRow, startCol + 2).Value = wsKpi.Cells(r, 6).Value
        ws.Cells(outRow, startCol + 3).Value = wsKpi.Cells(r, 7).Value
        ws.Cells(outRow, startCol + 2).NumberFormat = "0.0"
        ws.Cells(outRow, startCol + 3).NumberFormat = "0.0"

        Dim age As Double, sla As Double, ratio As Double
        age = NzNum(wsKpi.Cells(r, 6).Value)
        sla = NzNum(wsKpi.Cells(r, 7).Value)
        ratio = IIf(sla > 0, age / sla, 0)
        Dim col As Long
        If ratio >= 1 Then col = CLR_BAD ElseIf ratio >= 0.75 Then col = CLR_WARN Else col = CLR_GOOD
        ws.Cells(outRow, startCol + 4).Value = Format(ratio, "0%") & " of SLA"
        ws.Cells(outRow, startCol + 4).Interior.Color = col
        ws.Cells(outRow, startCol + 4).Font.Color = vbWhite
        ws.Cells(outRow, startCol + 4).Font.Bold = True
        ws.Cells(outRow, startCol + 4).HorizontalAlignment = xlCenter
        outRow = outRow + 1
    Next r

    ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(outRow - 1, startCol + 4)).Borders.LineStyle = xlContinuous
    ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(outRow - 1, startCol + 4)).Borders.Color = CLR_BORDER
End Sub

'==============================================================================
'                          3. INCIDENT ANALYTICS (War Room)
'==============================================================================
Public Sub BuildIncidentAnalytics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_INC)
    ResetSheetForDashboard ws
    PaintCanvas ws
    BuildHeaderBar ws, "INCIDENT WAR ROOM", _
        "Active incidents, escalations and root-cause hotlist"
    BuildNavigationRail ws, SHT_INC

    Dim cards As Variant
    cards = Array( _
        Array("Open", "OpenTickets", "0", CLR_WARN, ChrW(9888)), _
        Array("Critical Open", "CriticalOpen", "0", CLR_BAD, ChrW(9760)), _
        Array("Escalation %", "Escalation_Rate", "0.0%", CLR_BAD, ChrW(8607)), _
        Array("Repeat", "RepeatTickets", "0", CLR_PURPLE, ChrW(8634)), _
        Array("Aging > 3d", "AgingOver3d", "0", CLR_WARN, ChrW(9203)), _
        Array("Backlog %", "Backlog_Index", "0.0%", CLR_ACCENT2, ChrW(9776)))
    DrawCardRow ws, "inc", cards, 4, 6

    BuildSeverityStatusHeat ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 480, 250
    BuildAgingFunnel ws, _
        ws.Cells(11, 13).Left, ws.Cells(11, 13).Top, 380, 250, "Aging Funnel"
    BuildHourHeatChart ws, _
        ws.Cells(11, 22).Left, ws.Cells(11, 22).Top, 360, 250

    ' Open Critical drill (live console)
    BuildOpenIncidentsTable ws, 26, 2
End Sub

Private Sub BuildOpenIncidentsTable(ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)
    PaintSectionTitle ws, startRow, startCol, 22, "ACTIVE INCIDENTS  -  Open & Severity Critical/High"

    Dim arr As Variant: arr = ReadCleanArrayPub
    If IsEmpty(arr) Then Exit Sub

    Dim hdrRow As Long: hdrRow = startRow + 1
    Dim headers As Variant: headers = Array("ID", "Severity", "Title", "Reporter", "Location", "Reported", "Age (h)", "Status")
    Dim i As Long
    For i = 0 To UBound(headers): ws.Cells(hdrRow, startCol + i).Value = headers(i): Next i
    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(hdrRow, startCol + 7))
        .Font.Color = CLR_ACCENT
        .Font.Bold = True
        .Interior.Color = CLR_PANEL
    End With

    Dim outRow As Long: outRow = hdrRow + 1
    Dim r As Long
    For r = 2 To UBound(arr, 1)
        If CStr(arr(r, CD_NORM_STATUS)) = "Open" And _
           (CStr(arr(r, CD_NORM_SEVERITY)) = "Critical" Or CStr(arr(r, CD_NORM_SEVERITY)) = "High") Then
            ws.Cells(outRow, startCol).Value = arr(r, CD_ISSUE_ID)
            ws.Cells(outRow, startCol + 1).Value = arr(r, CD_NORM_SEVERITY)
            ws.Cells(outRow, startCol + 2).Value = arr(r, CD_ISSUE_TITLE)
            ws.Cells(outRow, startCol + 3).Value = arr(r, CD_REPORTER)
            ws.Cells(outRow, startCol + 4).Value = arr(r, CD_REP_LOC)
            ws.Cells(outRow, startCol + 5).Value = arr(r, CD_REPORT_DT)
            ws.Cells(outRow, startCol + 6).Value = arr(r, CD_AGE_HOURS)
            ws.Cells(outRow, startCol + 7).Value = arr(r, CD_NORM_STATUS)

            ws.Cells(outRow, startCol + 5).NumberFormat = "yyyy-mm-dd hh:mm"
            ws.Cells(outRow, startCol + 6).NumberFormat = "0.0"

            ' Severity pill
            ws.Cells(outRow, startCol + 1).Interior.Color = SeverityColor(CStr(arr(r, CD_NORM_SEVERITY)))
            ws.Cells(outRow, startCol + 1).Font.Color = vbWhite
            ws.Cells(outRow, startCol + 1).Font.Bold = True
            ws.Cells(outRow, startCol + 1).HorizontalAlignment = xlCenter
            outRow = outRow + 1
        End If
    Next r

    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(outRow - 1, startCol + 7))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = CLR_BORDER
        .Font.Size = 9
    End With
    ws.Cells(hdrRow, startCol).Resize(outRow - hdrRow, 1).ColumnWidth = 7
    ws.Cells(hdrRow, startCol + 1).ColumnWidth = 11
    ws.Cells(hdrRow, startCol + 2).ColumnWidth = 38
    ws.Cells(hdrRow, startCol + 3).ColumnWidth = 22
    ws.Cells(hdrRow, startCol + 4).ColumnWidth = 22
    ws.Cells(hdrRow, startCol + 5).ColumnWidth = 16
    ws.Cells(hdrRow, startCol + 6).ColumnWidth = 9
    ws.Cells(hdrRow, startCol + 7).ColumnWidth = 9
End Sub

'==============================================================================
'                          4. AGENT (RESOLVER) ANALYTICS
'==============================================================================
Public Sub BuildAgentAnalytics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_AGENT)
    ResetSheetForDashboard ws
    PaintCanvas ws
    BuildHeaderBar ws, "AGENT PRODUCTIVITY CENTER", _
        "Resolver workload, MTTR and efficiency leaderboard"
    BuildNavigationRail ws, SHT_AGENT

    Dim cards As Variant
    cards = Array( _
        Array("Closed", "ClosedTickets", "0", CLR_GOOD, ChrW(10003)), _
        Array("MTTR (h)", "MTTR_Hours", "0.0", CLR_ACCENT2, ChrW(8987)), _
        Array("Eff Score", "Agent_Efficiency", "0", CLR_ACCENT, ChrW(9889)), _
        Array("Auto Close %", "Auto_Closure_Rate", "0.0%", CLR_PURPLE, ChrW(8634)), _
        Array("Util %", "Support_Utilization", "0", CLR_ACCENT2, ChrW(9776)), _
        Array("Closure %", "ClosureRate", "0.0%", CLR_GOOD, ChrW(10004)))
    DrawCardRow ws, "agt", cards, 4, 6

    BuildResolverChart ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 480, 280
    BuildResolverEfficiencyChart ws, _
        ws.Cells(11, 13).Left, ws.Cells(11, 13).Top, 480, 280

    ' Resolver leaderboard table
    BuildResolverLeaderboard ws, 26, 2
End Sub

Private Sub BuildResolverLeaderboard(ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)
    PaintSectionTitle ws, startRow, startCol, 22, "RESOLVER LEADERBOARD"
    Dim r1 As Long, r2 As Long
    LocateBlock "AGG: Resolver MTTR", r1, r2
    If r1 = 0 Then Exit Sub
    Dim dm As Worksheet: Set dm = ThisWorkbook.Worksheets(SHT_MODEL)

    Dim hdrRow As Long: hdrRow = startRow + 1
    Dim headers As Variant: headers = Array("Resolver", "Tickets", "Avg MTTR (h)", "SLA Met %", "Efficiency", "Tier")
    Dim i As Long
    For i = 0 To UBound(headers): ws.Cells(hdrRow, startCol + i).Value = headers(i): Next i
    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(hdrRow, startCol + 5))
        .Font.Color = CLR_ACCENT: .Font.Bold = True: .Interior.Color = CLR_PANEL
    End With

    Dim outRow As Long: outRow = hdrRow + 1
    Dim r As Long, eff As Double, tier As String, col As Long
    For r = r1 + 1 To r2
        ws.Cells(outRow, startCol).Value = dm.Cells(r, 1).Value
        ws.Cells(outRow, startCol + 1).Value = dm.Cells(r, 2).Value
        ws.Cells(outRow, startCol + 2).Value = dm.Cells(r, 3).Value
        ws.Cells(outRow, startCol + 3).Value = dm.Cells(r, 4).Value
        ws.Cells(outRow, startCol + 4).Value = dm.Cells(r, 5).Value
        ws.Cells(outRow, startCol + 2).NumberFormat = "0.0"
        ws.Cells(outRow, startCol + 3).NumberFormat = "0.0%"
        ws.Cells(outRow, startCol + 4).NumberFormat = "0"
        eff = NzNum(dm.Cells(r, 5).Value)
        If eff >= 80 Then tier = "Star": col = CLR_GOOD _
        ElseIf eff >= 60 Then tier = "Solid": col = CLR_ACCENT _
        ElseIf eff >= 40 Then tier = "Watch": col = CLR_WARN _
        Else tier = "Coach": col = CLR_BAD
        ws.Cells(outRow, startCol + 5).Value = tier
        ws.Cells(outRow, startCol + 5).Interior.Color = col
        ws.Cells(outRow, startCol + 5).Font.Color = vbWhite
        ws.Cells(outRow, startCol + 5).Font.Bold = True
        ws.Cells(outRow, startCol + 5).HorizontalAlignment = xlCenter
        outRow = outRow + 1
    Next r

    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(outRow - 1, startCol + 5))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = CLR_BORDER
        .Font.Size = 9
    End With
End Sub

'==============================================================================
'                          5. APPLICATION HEALTH
'==============================================================================
Public Sub BuildApplicationAnalytics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_APP)
    ResetSheetForDashboard ws
    PaintCanvas ws
    BuildHeaderBar ws, "APPLICATION HEALTH CENTER", _
        "Issue type, location and equipment health monitoring"
    BuildNavigationRail ws, SHT_APP

    Dim cards As Variant
    cards = Array( _
        Array("Total", "TotalTickets", "0", CLR_ACCENT, ChrW(8505)), _
        Array("Critical %", "Critical_Failure_Rate", "0.0%", CLR_BAD, ChrW(9760)), _
        Array("Stability", "Stability_Score", "0", CLR_GOOD, ChrW(9776)), _
        Array("Risk", "Risk_Score", "0", CLR_BAD, ChrW(9889)), _
        Array("Downtime Risk", "Downtime_Risk", "0", CLR_WARN, ChrW(9203)), _
        Array("Impact", "Impact_Score", "0.0", CLR_PURPLE, ChrW(9889)))
    DrawCardRow ws, "app", cards, 4, 6

    BuildIssueTypeChart ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 380, 260
    BuildLocationVolumeChart ws, _
        ws.Cells(11, 11).Left, ws.Cells(11, 11).Top, 460, 260
    BuildLocationRiskChart ws, _
        ws.Cells(11, 22).Left, ws.Cells(11, 22).Top, 380, 260
End Sub

'==============================================================================
'                          6. RCA ANALYTICS (Root Cause)
'==============================================================================
Public Sub BuildRcaAnalytics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_RCA)
    ResetSheetForDashboard ws
    PaintCanvas ws
    BuildHeaderBar ws, "RCA ANALYTICS  -  Root Cause Tree", _
        "Auto-clustered themes with recurrence and recommendations"
    BuildNavigationRail ws, SHT_RCA

    Dim cards As Variant
    cards = Array( _
        Array("Repeat", "RepeatTickets", "0", CLR_PURPLE, ChrW(8634)), _
        Array("Repeat %", "Repeat_Rate", "0.0%", CLR_PURPLE, ChrW(8635)), _
        Array("Categories", "TotalTickets", "0", CLR_ACCENT, ChrW(9926)), _
        Array("Risk", "Risk_Score", "0", CLR_BAD, ChrW(9889)), _
        Array("Aging > 3d", "AgingOver3d", "0", CLR_WARN, ChrW(9203)), _
        Array("Critical Open", "CriticalOpen", "0", CLR_BAD, ChrW(9760)))
    DrawCardRow ws, "rca", cards, 4, 6

    BuildCategoryBars ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 480, 280, "Auto-Clustered Categories"
    BuildRcaTree ws, 11, 13

    ' Recommendations panel from KPI sheet
    BuildAiRecommendationsTable ws, 26, 2
End Sub

Private Sub BuildRcaTree(ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)
    ' Visual root-cause tree using shapes - parent "All Tickets" -> child categories
    Dim arr As Variant: arr = ReadCleanArrayPub
    If IsEmpty(arr) Then Exit Sub
    Dim cat As Object: Set cat = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To UBound(arr, 1)
        k = CStr(arr(r, CD_CATEGORY))
        If cat.Exists(k) Then cat(k) = cat(k) + 1 Else cat(k) = 1
    Next r
    Dim keys As Variant: keys = cat.keys
    SortDictDesc cat, keys

    PaintSectionTitle ws, startRow, startCol, 16, "ROOT CAUSE TREE  (auto-clustered)"

    Dim baseLeft As Double, baseTop As Double
    baseLeft = ws.Cells(startRow + 1, startCol).Left
    baseTop = ws.Cells(startRow + 1, startCol).Top

    ' Root node
    Dim root As Shape
    Set root = AddShape(ws, msoShapeRoundedRectangle, "rca_root", _
        baseLeft + 130, baseTop + 8, 140, 36)
    StyleShape root, CLR_ACCENT, 0
    root.Adjustments.Item(1) = 0.4
    root.TextFrame2.TextRange.Text = "ALL TICKETS  (" & UBound(arr, 1) - 1 & ")"
    With root.TextFrame2.TextRange.Font
        .Bold = msoTrue: .Size = 10: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_BG
    End With
    root.TextFrame2.HorizontalAnchor = msoAnchorCenter
    root.TextFrame2.VerticalAnchor = msoAnchorMiddle

    ' Children stacked vertically with connector lines
    Dim n As Long, top As Long
    top = WorksheetFunction.Min(8, UBound(keys))
    Dim childY As Double, childH As Double, gap As Double
    childH = 24: gap = 6
    For n = 0 To top
        childY = baseTop + 60 + n * (childH + gap)
        Dim child As Shape
        Set child = AddShape(ws, msoShapeRoundedRectangle, "rca_c" & n, _
            baseLeft + 60, childY, 280, childH)
        StyleShape child, CLR_PANEL_HI, 0
        child.Adjustments.Item(1) = 0.3
        child.TextFrame2.TextRange.Text = CStr(keys(n)) & "   " & ChrW(8226) & "   " & cat(keys(n)) & " tickets"
        With child.TextFrame2.TextRange.Font
            .Size = 9.5: .Name = "Segoe UI"
            .Fill.ForeColor.RGB = CLR_TEXT
        End With
        child.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

        ' Connector line
        Dim ln As Shape
        Set ln = ws.Shapes.AddConnector(msoConnectorElbow, _
            baseLeft + 200, baseTop + 44, _
            baseLeft + 60, childY + childH / 2)
        ln.name = "rca_l" & n
        ln.Line.ForeColor.RGB = CLR_BORDER
        ln.Line.Weight = 1.25
    Next n
End Sub

Private Sub BuildAiRecommendationsTable(ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)
    PaintSectionTitle ws, startRow, startCol, 22, "AI RECOMMENDATIONS"
    Dim wsKpi As Worksheet: Set wsKpi = ThisWorkbook.Worksheets(SHT_KPI)
    Dim r As Long, lr As Long, foundAt As Long
    lr = wsKpi.Cells(wsKpi.Rows.count, 4).End(xlUp).Row
    For r = 1 To lr
        If CStr(wsKpi.Cells(r, 4).Value) = "AI RECOMMENDATIONS" Then foundAt = r + 2: Exit For
    Next r
    If foundAt = 0 Then Exit Sub

    Dim hdrRow As Long: hdrRow = startRow + 1
    Dim headers As Variant: headers = Array("Theme", "Insight", "Recommended Action")
    Dim i As Long
    For i = 0 To UBound(headers): ws.Cells(hdrRow, startCol + i).Value = headers(i): Next i
    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(hdrRow, startCol + 2))
        .Font.Color = CLR_ACCENT: .Font.Bold = True: .Interior.Color = CLR_PANEL
    End With

    Dim outRow As Long: outRow = hdrRow + 1
    For r = foundAt To lr
        Dim v As String: v = CStr(wsKpi.Cells(r, 4).Value)
        If Len(v) = 0 Then Exit For
        If UCase$(v) = "PREDICTED SLA BREACHES" Then Exit For
        ws.Cells(outRow, startCol).Value = wsKpi.Cells(r, 4).Value
        ws.Cells(outRow, startCol + 1).Value = wsKpi.Cells(r, 5).Value
        ws.Cells(outRow, startCol + 2).Value = wsKpi.Cells(r, 6).Value
        outRow = outRow + 1
    Next r

    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(outRow - 1, startCol + 2))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = CLR_BORDER
        .Font.Size = 9
        .WrapText = True
        .VerticalAlignment = xlTop
    End With
    ws.Cells(hdrRow, startCol).ColumnWidth = 22
    ws.Cells(hdrRow, startCol + 1).ColumnWidth = 38
    ws.Cells(hdrRow, startCol + 2).ColumnWidth = 50
End Sub

'==============================================================================
'                          7. FORECAST ANALYTICS
'==============================================================================
Public Sub BuildForecastAnalytics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_FORECAST)
    ResetSheetForDashboard ws
    PaintCanvas ws
    BuildHeaderBar ws, "OPERATIONAL FORECAST CENTER", _
        "3-day projected demand, peak hours, and capacity planning"
    BuildNavigationRail ws, SHT_FORECAST

    Dim cards As Variant
    Dim peak As Variant: peak = GetKpi("Peak_Hour")
    cards = Array( _
        Array("Trend WoW %", "Trend_Velocity", "0.0", CLR_ACCENT, ChrW(8599)), _
        Array("Total", "TotalTickets", "0", CLR_ACCENT2, ChrW(8505)), _
        Array("Open", "OpenTickets", "0", CLR_WARN, ChrW(9888)), _
        Array("MTTR (h)", "MTTR_Hours", "0.0", CLR_ACCENT2, ChrW(8987)), _
        Array("Aging > 3d", "AgingOver3d", "0", CLR_WARN, ChrW(9203)), _
        Array("Util %", "Support_Utilization", "0", CLR_PURPLE, ChrW(9776)))
    DrawCardRow ws, "fc", cards, 4, 6

    ' Peak hour pill
    Dim pill As Shape
    Set pill = AddShape(ws, msoShapeRoundedRectangle, "fc_peak", _
        ws.Cells(8, 2).Left, ws.Cells(8, 2).Top, 380, 32)
    StyleShape pill, CLR_PANEL_HI, 0
    pill.TextFrame2.TextRange.Text = "PEAK INTAKE WINDOW:  " & CStr(peak)
    With pill.TextFrame2.TextRange.Font
        .Bold = msoTrue: .Size = 10: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_ACCENT
    End With
    pill.TextFrame2.HorizontalAnchor = msoAnchorCenter
    pill.TextFrame2.VerticalAnchor = msoAnchorMiddle

    BuildForecastChart ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 540, 280
    BuildHourHeatChart ws, _
        ws.Cells(11, 14).Left, ws.Cells(11, 14).Top, 460, 280
    BuildDowChart ws, _
        ws.Cells(11, 24).Left, ws.Cells(11, 24).Top, 380, 280

    ' Hour x DOW heatmap (cell-based)
    PaintSectionTitle ws, 26, 2, 28, "HOUR x DAY-OF-WEEK INTAKE HEATMAP"
    BuildHourDowHeatmap ws, 27, 2
End Sub

'==============================================================================
'                          8. RISK ANALYTICS
'==============================================================================
Public Sub BuildRiskAnalytics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_RISK)
    ResetSheetForDashboard ws
    PaintCanvas ws
    BuildHeaderBar ws, "RISK MONITORING CENTER", _
        "Composite risk, downtime, and high-impact site watch-list"
    BuildNavigationRail ws, SHT_RISK

    Dim cards As Variant
    cards = Array( _
        Array("Risk Score", "Risk_Score", "0", CLR_BAD, ChrW(9889)), _
        Array("Downtime", "Downtime_Risk", "0", CLR_BAD, ChrW(9203)), _
        Array("Aging Risk", "Aging_Risk", "0", CLR_WARN, ChrW(9203)), _
        Array("Critical %", "Critical_Failure_Rate", "0.0%", CLR_BAD, ChrW(9760)), _
        Array("Priority Heat", "Priority_Heat", "0.0", CLR_WARN, ChrW(9889)), _
        Array("Stability", "Stability_Score", "0", CLR_GOOD, ChrW(9776)))
    DrawCardRow ws, "rsk", cards, 4, 6

    BuildLocationRiskChart ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 460, 280
    BuildResolverEfficiencyChart ws, _
        ws.Cells(11, 13).Left, ws.Cells(11, 13).Top, 460, 280
    BuildOpsRadar ws, _
        ws.Cells(11, 24).Left, ws.Cells(11, 24).Top, 380, 280

    ' Risk matrix table
    BuildRiskMatrix ws, 26, 2
End Sub

Private Sub BuildRiskMatrix(ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)
    PaintSectionTitle ws, startRow, startCol, 24, "INCIDENT RISK MATRIX  -  Severity x Status with risk colour-scale"

    Dim hdrRow As Long: hdrRow = startRow + 1
    Dim sevs As Variant: sevs = Array("Critical", "High", "Medium", "Low", "Unspecified")
    Dim stats As Variant: stats = Array("Open", "Closed")

    ws.Cells(hdrRow, startCol).Value = "Severity \ Status"
    Dim i As Long, j As Long
    For j = 0 To UBound(stats)
        ws.Cells(hdrRow, startCol + 1 + j).Value = stats(j)
    Next j
    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(hdrRow, startCol + 2))
        .Font.Color = CLR_ACCENT: .Font.Bold = True: .Interior.Color = CLR_PANEL
    End With

    Dim arr As Variant: arr = ReadCleanArrayPub
    If IsEmpty(arr) Then Exit Sub
    Dim r As Long, sv As String, st As String, key As String
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(arr, 1)
        sv = CStr(arr(r, CD_NORM_SEVERITY))
        st = CStr(arr(r, CD_NORM_STATUS))
        key = sv & "|" & st
        If d.Exists(key) Then d(key) = d(key) + 1 Else d(key) = 1
    Next r

    Dim outRow As Long: outRow = hdrRow + 1
    For i = 0 To UBound(sevs)
        ws.Cells(outRow, startCol).Value = sevs(i)
        ws.Cells(outRow, startCol).Interior.Color = SeverityColor(CStr(sevs(i)))
        ws.Cells(outRow, startCol).Font.Color = vbWhite
        ws.Cells(outRow, startCol).Font.Bold = True
        For j = 0 To UBound(stats)
            key = CStr(sevs(i)) & "|" & CStr(stats(j))
            ws.Cells(outRow, startCol + 1 + j).Value = IIf(d.Exists(key), d(key), 0)
            ws.Cells(outRow, startCol + 1 + j).HorizontalAlignment = xlCenter
        Next j
        outRow = outRow + 1
    Next i

    Dim heat As Range
    Set heat = ws.Range(ws.Cells(hdrRow + 1, startCol + 1), ws.Cells(outRow - 1, startCol + 2))
    heat.FormatConditions.Delete
    Dim cs As ColorScale
    Set cs = heat.FormatConditions.AddColorScale(ColorScaleType:=3)
    cs.ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    cs.ColorScaleCriteria(1).FormatColor.Color = CLR_PANEL
    cs.ColorScaleCriteria(2).Type = xlConditionValuePercentile
    cs.ColorScaleCriteria(2).Value = 50
    cs.ColorScaleCriteria(2).FormatColor.Color = CLR_WARN
    cs.ColorScaleCriteria(3).Type = xlConditionValueHighestValue
    cs.ColorScaleCriteria(3).FormatColor.Color = CLR_BAD

    With ws.Range(ws.Cells(hdrRow, startCol), ws.Cells(outRow - 1, startCol + 2))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = CLR_BORDER
        .Font.Size = 10
        .RowHeight = 22
    End With
End Sub

'==============================================================================
'                          SHARED HELPERS
'==============================================================================
Private Sub DrawCardRow(ws As Worksheet, ByVal idPrefix As String, ByVal cards As Variant, _
        ByVal rowAnchor As Long, ByVal n As Long)
    Dim startLeft As Double, top As Double, gap As Double
    startLeft = ws.Cells(rowAnchor, 2).Left
    top = ws.Cells(rowAnchor, 2).Top
    gap = 8
    Dim totalW As Double: totalW = ws.Range("B" & rowAnchor & ":AB" & rowAnchor).Width
    Dim cardW As Double: cardW = (totalW - gap * (n - 1)) / n
    Dim i As Long
    For i = 0 To n - 1
        Dim arr As Variant: arr = cards(i)
        DrawKpiCard ws, _
            "kpi_" & idPrefix & "_" & i, _
            startLeft + i * (cardW + gap), top, cardW, KPI_CARD_H, _
            CStr(arr(0)), CStr(arr(1)), CStr(arr(2)), CLng(arr(3)), CStr(arr(4))
    Next i
End Sub

Private Sub PaintSectionTitle(ws As Worksheet, ByVal r As Long, ByVal c As Long, _
        ByVal cols As Long, ByVal txt As String)
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(r, c), ws.Cells(r, c + cols - 1))
    rng.Merge
    rng.Value = txt
    rng.Interior.Color = CLR_PANEL
    rng.Font.Color = CLR_ACCENT
    rng.Font.Bold = True
    rng.Font.Size = 11
    rng.Font.Name = "Segoe UI Semibold"
    rng.HorizontalAlignment = xlLeft
    rng.RowHeight = 24
End Sub



'==============================================================================
' === SECTION: M07_Interaction ===
'==============================================================================

'==============================================================================
' MODULE      : M07_Interaction
' DESCRIPTION : Interaction layer.  Navigation rail, command buttons, smart
'               search, drill-through, theme switching, PDF export, email
'               distribution, settings/logs scaffolding and login auth.
'==============================================================================

'==============================================================================
'                          NAVIGATION RAIL
'==============================================================================
Public Sub WireNavigationEverywhere()
    Dim sheets As Variant, i As Long
    sheets = Array(SHT_DASH, SHT_EXEC, SHT_SLA, SHT_INC, SHT_AGENT, _
                   SHT_APP, SHT_RCA, SHT_FORECAST, SHT_RISK, SHT_SETTINGS)
    For i = LBound(sheets) To UBound(sheets)
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(sheets(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then BuildNavigationRail ws, CStr(sheets(i))
        Set ws = Nothing
    Next i
End Sub

Public Sub BuildNavigationRail(ws As Worksheet, ByVal currentSheet As String)
    ' Remove existing nav shapes
    DeleteShapesByPrefix ws, "nav_"

    ' Side rail panel
    Dim panel As Shape
    Set panel = AddShape(ws, msoShapeRectangle, "nav_panel", _
        ws.Range("A1").Left, ws.Range("A1").Top + HEADER_HEIGHT + 4, _
        NAV_WIDTH - 6, 480)
    StyleShape panel, CLR_PANEL, 0
    AddShadow panel

    ' Section title
    Dim ttl As Shape
    Set ttl = AddShape(ws, msoShapeRectangle, "nav_ttl", _
        ws.Range("A1").Left + 8, ws.Range("A1").Top + HEADER_HEIGHT + 12, NAV_WIDTH - 22, 22)
    StyleShape ttl, CLR_PANEL, 0
    ttl.TextFrame2.TextRange.Text = "MODULES"
    With ttl.TextFrame2.TextRange.Font
        .Size = 9: .Bold = msoTrue: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_MUTED
    End With

    ' Nav entries
    Dim items As Variant
    items = Array( _
        Array(SHT_DASH, "Command Center", ChrW(9632)), _
        Array(SHT_EXEC, "Executive View", ChrW(9650)), _
        Array(SHT_SLA, "SLA Intelligence", ChrW(10003)), _
        Array(SHT_INC, "Incident War Room", ChrW(9888)), _
        Array(SHT_AGENT, "Agent Productivity", ChrW(9786)), _
        Array(SHT_APP, "Application Health", ChrW(9926)), _
        Array(SHT_RCA, "RCA Analytics", ChrW(8634)), _
        Array(SHT_FORECAST, "Forecast Center", ChrW(8599)), _
        Array(SHT_RISK, "Risk Monitoring", ChrW(9889)), _
        Array(SHT_SETTINGS, "Settings", ChrW(9881)))

    Dim i As Long, top As Double, h As Double, leftPad As Double
    leftPad = ws.Range("A1").Left + 8
    top = ws.Range("A1").Top + HEADER_HEIGHT + 40
    h = 30
    For i = LBound(items) To UBound(items)
        Dim item As Variant: item = items(i)
        DrawNavItem ws, "nav_i" & i, leftPad, top + i * (h + 4), NAV_WIDTH - 22, h, _
            CStr(item(0)), CStr(item(1)), CStr(item(2)), _
            (StrComp(CStr(item(0)), currentSheet, vbTextCompare) = 0)
    Next i
End Sub

Private Sub DrawNavItem(ws As Worksheet, ByVal name As String, _
        ByVal lft As Double, ByVal tp As Double, _
        ByVal w As Double, ByVal h As Double, _
        ByVal targetSheet As String, ByVal label As String, _
        ByVal glyph As String, ByVal isActive As Boolean)

    Dim card As Shape
    Set card = AddShape(ws, msoShapeRoundedRectangle, name, lft, tp, w, h)
    card.Adjustments.Item(1) = 0.3
    If isActive Then
        StyleShape card, CLR_ACCENT, 0
    Else
        StyleShape card, CLR_PANEL_HI, 0
    End If

    card.TextFrame2.TextRange.Text = "  " & glyph & "   " & label
    With card.TextFrame2.TextRange.Font
        .Size = 10: .Bold = msoTrue: .Name = "Segoe UI"
        If isActive Then .Fill.ForeColor.RGB = CLR_BG Else .Fill.ForeColor.RGB = CLR_TEXT
    End With
    card.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    card.TextFrame2.VerticalAnchor = msoAnchorMiddle
    card.OnAction = "'NavigateTo """ & targetSheet & """'"
End Sub

Public Sub NavigateTo(ByVal sheetName As String)
    On Error Resume Next
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub
    If ws.Visible <> xlSheetVisible Then ws.Visible = xlSheetVisible
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    Range("A1").Select
    LogInfo "NAV", "Navigated to " & sheetName
End Sub

'==============================================================================
'                          COMMAND BUTTON FACTORY
'==============================================================================
Public Sub AddCommandButton(ws As Worksheet, ByVal name As String, _
        ByVal lft As Double, ByVal tp As Double, _
        ByVal w As Double, ByVal h As Double, _
        ByVal label As String, ByVal macro As String, _
        ByVal fillCol As Long, ByVal textCol As Long)

    Dim btn As Shape
    Set btn = AddShape(ws, msoShapeRoundedRectangle, name, lft, tp, w, h)
    btn.Adjustments.Item(1) = 0.3
    StyleShape btn, fillCol, 0
    AddShadow btn
    btn.TextFrame2.TextRange.Text = label
    With btn.TextFrame2.TextRange.Font
        .Bold = msoTrue: .Size = 10: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = textCol
    End With
    btn.TextFrame2.HorizontalAnchor = msoAnchorCenter
    btn.TextFrame2.VerticalAnchor = msoAnchorMiddle
    btn.OnAction = macro
End Sub

'==============================================================================
'                          SMART SEARCH (drill-through)
'==============================================================================
Public Sub SmartSearch()
    Dim term As String
    term = InputBox("Search tickets by ID, title, reporter, location or category:", _
                    "Smart Search  -  " & APP_SHORT, "")
    If Len(Trim$(term)) = 0 Then Exit Sub
    DrillThrough term
End Sub

Public Sub DrillThrough(ByVal term As String)
    On Error GoTo Fail
    Application.ScreenUpdating = False
    Dim arr As Variant: arr = ReadCleanArrayPub
    If IsEmpty(arr) Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Search_Results")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = "Search_Results"
    End If
    ResetSheetForDashboard ws
    PaintCanvas ws
    BuildHeaderBar ws, "SEARCH RESULTS  -  """ & term & """", _
        "Filtered tickets matching the search expression"
    BuildNavigationRail ws, "Search_Results"

    ' Header
    Dim hdrRow As Long: hdrRow = 5
    Dim col As Long: col = 2
    Dim headers As Variant
    headers = Array("ID", "Sev", "Status", "Title", "Reporter", "Location", _
                    "Type", "Category", "Reported", "Age (h)", "SLA Status", "Resolver")
    Dim i As Long
    For i = 0 To UBound(headers): ws.Cells(hdrRow, col + i).Value = headers(i): Next i
    With ws.Range(ws.Cells(hdrRow, col), ws.Cells(hdrRow, col + UBound(headers)))
        .Interior.Color = CLR_PANEL
        .Font.Color = CLR_ACCENT: .Font.Bold = True
    End With

    Dim outRow As Long: outRow = hdrRow + 1
    Dim r As Long, hay As String
    Dim t As String: t = LCase$(term)
    For r = 2 To UBound(arr, 1)
        hay = LCase$(CStr(arr(r, CD_ISSUE_ID)) & " " & arr(r, CD_ISSUE_TITLE) & " " & _
                     arr(r, CD_REPORTER) & " " & arr(r, CD_REP_LOC) & " " & _
                     arr(r, CD_CATEGORY) & " " & arr(r, CD_NORM_ISSUE_TYPE) & " " & _
                     arr(r, CD_RESOLVER) & " " & arr(r, CD_NORM_SEVERITY) & " " & _
                     arr(r, CD_NORM_STATUS))
        If InStr(hay, t) > 0 Then
            ws.Cells(outRow, col).Value = arr(r, CD_ISSUE_ID)
            ws.Cells(outRow, col + 1).Value = arr(r, CD_NORM_SEVERITY)
            ws.Cells(outRow, col + 2).Value = arr(r, CD_NORM_STATUS)
            ws.Cells(outRow, col + 3).Value = arr(r, CD_ISSUE_TITLE)
            ws.Cells(outRow, col + 4).Value = arr(r, CD_REPORTER)
            ws.Cells(outRow, col + 5).Value = arr(r, CD_REP_LOC)
            ws.Cells(outRow, col + 6).Value = arr(r, CD_NORM_ISSUE_TYPE)
            ws.Cells(outRow, col + 7).Value = arr(r, CD_CATEGORY)
            ws.Cells(outRow, col + 8).Value = arr(r, CD_REPORT_DT)
            ws.Cells(outRow, col + 9).Value = arr(r, CD_AGE_HOURS)
            ws.Cells(outRow, col + 10).Value = arr(r, CD_SLA_STATUS)
            ws.Cells(outRow, col + 11).Value = arr(r, CD_RESOLVER)

            ws.Cells(outRow, col + 8).NumberFormat = "yyyy-mm-dd hh:mm"
            ws.Cells(outRow, col + 9).NumberFormat = "0.0"

            ' Severity pill
            ws.Cells(outRow, col + 1).Interior.Color = SeverityColor(CStr(arr(r, CD_NORM_SEVERITY)))
            ws.Cells(outRow, col + 1).Font.Color = vbWhite
            ws.Cells(outRow, col + 1).Font.Bold = True
            ws.Cells(outRow, col + 1).HorizontalAlignment = xlCenter
            ' Status pill
            If CStr(arr(r, CD_NORM_STATUS)) = "Open" Then
                ws.Cells(outRow, col + 2).Interior.Color = CLR_WARN
            Else
                ws.Cells(outRow, col + 2).Interior.Color = CLR_GOOD
            End If
            ws.Cells(outRow, col + 2).Font.Color = vbWhite
            ws.Cells(outRow, col + 2).Font.Bold = True
            ws.Cells(outRow, col + 2).HorizontalAlignment = xlCenter
            outRow = outRow + 1
        End If
    Next r

    ' Frame
    With ws.Range(ws.Cells(hdrRow, col), ws.Cells(WorksheetFunction.Max(outRow - 1, hdrRow + 1), col + UBound(headers)))
        .Borders.LineStyle = xlContinuous
        .Borders.Color = CLR_BORDER
        .Font.Size = 9
        .Font.Color = CLR_TEXT
    End With
    AutoSizeSearchCols ws

    ' Result count pill
    Dim pill As Shape
    Set pill = AddShape(ws, msoShapeRoundedRectangle, "srch_count", _
        ws.Cells(4, 2).Left, ws.Cells(4, 2).Top, 220, 22)
    StyleShape pill, CLR_PANEL_HI, 0
    pill.TextFrame2.TextRange.Text = "  " & ChrW(128269) & "  " & (outRow - hdrRow - 1) & " match(es)"
    With pill.TextFrame2.TextRange.Font
        .Size = 10: .Bold = msoTrue: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_ACCENT
    End With

    ws.Activate
    Range("A1").Select
    Application.ScreenUpdating = True
    Exit Sub
Fail:
    Application.ScreenUpdating = True
    LogError "SEARCH", Err.Number, Err.Description
End Sub

Private Sub AutoSizeSearchCols(ws As Worksheet)
    Dim widths As Variant: widths = Array(7, 11, 11, 38, 22, 22, 14, 22, 16, 9, 13, 24)
    Dim i As Long
    For i = 0 To UBound(widths)
        ws.Cells(1, 2 + i).ColumnWidth = CDbl(widths(i))
    Next i
End Sub

'==============================================================================
'                          DRILL-THROUGH FROM KPI CARD CLICKS
'==============================================================================
Public Sub DrillCriticalOpen()
    DrillThrough "Critical"
End Sub
Public Sub DrillBreached()
    DrillThrough "Breached"
End Sub
Public Sub DrillRepeat()
    DrillThrough "Repeat"
End Sub

'==============================================================================
'                          PDF EXPORT
'==============================================================================
Public Sub ExportDashboardPdf()
    On Error GoTo Fail
    PerfBegin
    Dim path As String
    path = ThisWorkbook.Path
    If Len(path) = 0 Then path = Environ$("USERPROFILE") & "\Documents"
    Dim fname As String
    fname = path & Application.PathSeparator & APP_SHORT & "_Dashboard_" & _
            Format(Now, "yyyymmdd_hhnnss") & ".pdf"

    ' Print only the visible analytics sheets
    Dim sheetsToExport As Variant, i As Long
    sheetsToExport = Array(SHT_DASH, SHT_EXEC, SHT_SLA, SHT_INC, SHT_AGENT, _
                           SHT_APP, SHT_RCA, SHT_FORECAST, SHT_RISK)
    Dim names() As String
    ReDim names(0 To UBound(sheetsToExport))
    Dim n As Long: n = 0
    For i = 0 To UBound(sheetsToExport)
        On Error Resume Next
        Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(CStr(sheetsToExport(i)))
        On Error GoTo Fail
        If Not ws Is Nothing Then
            ' Set page setup for nice export
            With ws.PageSetup
                .Orientation = xlLandscape
                .Zoom = False
                .FitToPagesWide = 1
                .FitToPagesTall = 1
                .PrintArea = ""
                .CenterHorizontally = True
                .CenterVertically = True
            End With
            names(n) = ws.name: n = n + 1
        End If
        Set ws = Nothing
    Next i
    ReDim Preserve names(0 To n - 1)

    ThisWorkbook.Worksheets(names).Select
    ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fname, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, _
        IgnorePrintAreas:=True, OpenAfterPublish:=True
    ' Re-select dashboard
    ThisWorkbook.Worksheets(SHT_DASH).Activate
    PerfRestore
    LogInfo "EXPORT", "Exported PDF -> " & fname
    MsgBox "Exported PDF:" & vbCrLf & fname, vbInformation, APP_SHORT
    Exit Sub
Fail:
    PerfRestore
    LogError "EXPORT", Err.Number, Err.Description
    MsgBox "PDF export failed: " & Err.Description, vbExclamation, APP_SHORT
End Sub

'==============================================================================
'                          EMAIL MIS REPORT
'==============================================================================
Public Sub EmailMisReport()
    On Error GoTo Fail
    Dim recipients As String, subj As String, body As String
    Dim wsCfg As Worksheet
    Set wsCfg = ThisWorkbook.Worksheets(SHT_HIDDEN)
    recipients = CStr(wsCfg.Range("B5").Value)
    If Len(Trim$(recipients)) = 0 Then
        recipients = InputBox("Recipient email(s), separated by ;", "Email MIS  -  " & APP_SHORT, _
                              "ops@example.com")
    End If
    If Len(Trim$(recipients)) = 0 Then Exit Sub

    subj = APP_NAME & "  -  Daily MIS  -  " & Format(Now, "dd-mmm-yyyy")
    body = BuildMisHtml()

    Dim ol As Object, mail As Object
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo Fail
    If ol Is Nothing Then
        ' Fallback - mailto: link
        Dim url As String
        url = "mailto:" & recipients & "?subject=" & EncodeUrl(subj) & "&body=" & _
              EncodeUrl("Open the attached workbook for full dashboard.")
        ThisWorkbook.FollowHyperlink url
        LogInfo "EMAIL", "Opened mailto fallback to " & recipients
        Exit Sub
    End If
    Set mail = ol.CreateItem(0)
    With mail
        .To = recipients
        .Subject = subj
        .HTMLBody = body
        .Display
    End With
    LogInfo "EMAIL", "Composed MIS email to " & recipients
    Exit Sub
Fail:
    LogError "EMAIL", Err.Number, Err.Description
    MsgBox "Email failed: " & Err.Description, vbExclamation, APP_SHORT
End Sub

Private Function BuildMisHtml() As String
    Dim h As String
    h = "<div style=""font-family:Segoe UI,Arial;color:#222;font-size:14px"">"
    h = h & "<h2 style=""color:#0078D4;margin-bottom:4px"">" & APP_NAME & "</h2>"
    h = h & "<p style=""color:#666;margin-top:0"">Daily MIS - " & Format(Now, "dd-mmm-yyyy hh:mm") & "</p>"
    h = h & "<table cellpadding=""8"" cellspacing=""0"" style=""border-collapse:collapse;width:100%;border:1px solid #ddd"">"
    h = h & "<tr style=""background:#0F1724;color:#00E5FF""><th align=""left"">KPI</th><th align=""right"">Value</th></tr>"
    h = h & MisRow("Total Tickets", FmtKpi("TotalTickets", "0"))
    h = h & MisRow("Open Tickets", FmtKpi("OpenTickets", "0"))
    h = h & MisRow("Critical Open", FmtKpi("CriticalOpen", "0"))
    h = h & MisRow("SLA Compliance", FmtKpi("SLA_Compliance", "0.0%"))
    h = h & MisRow("MTTR (h)", FmtKpi("MTTR_Hours", "0.00"))
    h = h & MisRow("Backlog Index", FmtKpi("Backlog_Index", "0.0%"))
    h = h & MisRow("Risk Score", FmtKpi("Risk_Score", "0.0"))
    h = h & MisRow("Ops Health", FmtKpi("OpsHealth_Score", "0.0"))
    h = h & MisRow("CSAT Index", FmtKpi("Customer_Satisfaction", "0.0"))
    h = h & MisRow("Trend WoW %", FmtKpi("Trend_Velocity", "0.0"))
    h = h & "</table>"
    h = h & "<p style=""color:#666;font-size:12px;margin-top:16px"">Generated by " & APP_SHORT & " v" & APP_VERSION & "</p>"
    h = h & "</div>"
    BuildMisHtml = h
End Function

Private Function MisRow(ByVal k As String, ByVal v As String) As String
    MisRow = "<tr><td style=""border-bottom:1px solid #eee"">" & k & _
            "</td><td align=""right"" style=""border-bottom:1px solid #eee;font-weight:600"">" & v & "</td></tr>"
End Function

Private Function EncodeUrl(ByVal s As String) As String
    Dim out As String, i As Long, c As String, code As Long
    For i = 1 To Len(s)
        c = Mid$(s, i, 1)
        code = AscW(c)
        If (code >= 48 And code <= 57) Or (code >= 65 And code <= 90) Or _
           (code >= 97 And code <= 122) Or InStr("-_.~", c) > 0 Then
            out = out & c
        ElseIf c = " " Then
            out = out & "%20"
        Else
            out = out & "%" & Right$("00" & Hex(code), 2)
        End If
    Next i
    EncodeUrl = out
End Function

'==============================================================================
'                          THEME SWITCHER
'==============================================================================
Public Sub ToggleTheme()
    On Error Resume Next
    Dim hidden As Worksheet: Set hidden = ThisWorkbook.Worksheets(SHT_HIDDEN)
    On Error GoTo 0
    Dim curTheme As String
    curTheme = CStr(hidden.Range("B2").Value)
    If curTheme = "Light" Then
        hidden.Range("B2").Value = "Dark"
    Else
        hidden.Range("B2").Value = "Light"
    End If
    ApplyTheme CStr(hidden.Range("B2").Value)
End Sub

Public Sub ApplyTheme(ByVal mode As String)
    ' This is intentionally lightweight - flips background of analytics sheets and
    ' rebuilds the dashboard.  Keeps the chart/shape engines untouched (they will
    ' pick up the new palette next BuildEnterpriseDashboard run).
    Dim sheets As Variant: sheets = Array(SHT_DASH, SHT_EXEC, SHT_SLA, SHT_INC, _
        SHT_AGENT, SHT_APP, SHT_RCA, SHT_FORECAST, SHT_RISK, SHT_SETTINGS)
    Dim i As Long, ws As Worksheet
    Dim bg As Long
    If LCase$(mode) = "light" Then bg = RGB(245, 247, 250) Else bg = CLR_BG
    For i = 0 To UBound(sheets)
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(sheets(i)))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.Cells.Interior.Color = bg
        End If
        Set ws = Nothing
    Next i
    LogInfo "THEME", "Theme switched to " & mode
End Sub

'==============================================================================
'                          LOGIN AUTHENTICATION
'==============================================================================
Public Sub PromptLogin()
    On Error GoTo Fail
    Dim u As String, p As String
    Dim hidden As Worksheet: Set hidden = ThisWorkbook.Worksheets(SHT_HIDDEN)
    Dim adminUser As String, adminPass As String
    adminUser = CStr(hidden.Range("B3").Value)
    adminPass = CStr(hidden.Range("B4").Value)
    If Len(adminUser) = 0 Then adminUser = "admin"
    If Len(adminPass) = 0 Then adminPass = "admin"

    u = InputBox("Username:", APP_SHORT & "  -  Sign In")
    If Len(u) = 0 Then Exit Sub
    p = InputBox("Password:", APP_SHORT & "  -  Sign In")
    If u = adminUser And p = adminPass Then
        hidden.Range("B6").Value = u
        hidden.Range("B7").Value = "Admin"
        hidden.Range("B8").Value = Now
        UpdateUserPill u, "Admin"
        LogInfo "AUTH", "Login success: " & u
    ElseIf Len(p) > 0 Then
        hidden.Range("B6").Value = u
        hidden.Range("B7").Value = "Viewer"
        hidden.Range("B8").Value = Now
        UpdateUserPill u, "Viewer"
        LogInfo "AUTH", "Viewer login: " & u
    Else
        LogWarn "AUTH", "Cancelled login"
    End If
    Exit Sub
Fail:
    LogError "AUTH", Err.Number, Err.Description
End Sub

Private Sub UpdateUserPill(ByVal u As String, ByVal role As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_DASH)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Dim sh As Shape
    On Error Resume Next
    Set sh = ws.Shapes("hdr_pill")
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub
    sh.TextFrame2.TextRange.Text = ChrW(9679) & "  " & UCase$(role) & "  |  " & u
    sh.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(role = "Admin", CLR_GOOD, CLR_ACCENT2)
End Sub

'==============================================================================
'                          SETTINGS / LOGS / HIDDEN
'==============================================================================
Public Sub BuildSettingsSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_SETTINGS)
    ResetSheetForDashboard ws
    PaintCanvas ws
    BuildHeaderBar ws, "SETTINGS  &  ADMINISTRATION", _
        "SLA targets, theme, security, refresh and backup controls"
    BuildNavigationRail ws, SHT_SETTINGS

    ' Section: SLA Targets
    DrawSettingsCard ws, "SLA TARGETS (hours)", 4, 2, _
        Array(Array("Critical", SLA_CRITICAL_H), _
              Array("High", SLA_HIGH_H), _
              Array("Medium", SLA_MEDIUM_H), _
              Array("Low", SLA_LOW_H), _
              Array("Unspecified", SLA_UNSPEC_H))

    ' Section: Theme & Security
    DrawSettingsCard ws, "THEME  &  SECURITY", 4, 10, _
        Array(Array("Theme", "Dark"), _
              Array("Admin User", "admin"), _
              Array("Last Login", Format(Now, "yyyy-mm-dd hh:mm")), _
              Array("Build Date", Format(Now, "yyyy-mm-dd")), _
              Array("Version", APP_VERSION))

    ' Section: Counts (mirrored from KPI_Engine)
    DrawSettingsCard ws, "DATASET STATUS", 4, 18, _
        Array(Array("Total Tickets", GetKpi("TotalTickets")), _
              Array("Open", GetKpi("OpenTickets")), _
              Array("Closed", GetKpi("ClosedTickets")), _
              Array("Last Refresh", Format(GetKpi("LastRefreshed"), "yyyy-mm-dd hh:mm")), _
              Array("Source", "Apps.csv"))

    ' Action buttons
    Dim btnLeft As Double, btnTop As Double
    btnLeft = ws.Cells(20, 2).Left
    btnTop = ws.Cells(20, 2).Top
    AddCommandButton ws, "set_btn_refresh", btnLeft, btnTop, 150, 32, _
        ChrW(8635) & "  REFRESH ALL", "RefreshAll", CLR_ACCENT, CLR_BG
    AddCommandButton ws, "set_btn_rebuild", btnLeft + 160, btnTop, 150, 32, _
        ChrW(9881) & "  REBUILD ALL", "BuildEnterpriseDashboard", CLR_PURPLE, vbWhite
    AddCommandButton ws, "set_btn_pdf", btnLeft + 320, btnTop, 150, 32, _
        ChrW(8595) & "  EXPORT PDF", "ExportDashboardPdf", CLR_ACCENT2, CLR_BG
    AddCommandButton ws, "set_btn_email", btnLeft + 480, btnTop, 150, 32, _
        ChrW(9993) & "  EMAIL MIS", "EmailMisReport", CLR_GOOD, vbWhite
    AddCommandButton ws, "set_btn_login", btnLeft + 640, btnTop, 150, 32, _
        ChrW(128272) & "  SIGN IN", "PromptLogin", CLR_PANEL_HI, CLR_TEXT
    AddCommandButton ws, "set_btn_backup", btnLeft, btnTop + 40, 150, 32, _
        ChrW(128190) & "  BACKUP", "BackupWorkbook", CLR_PANEL_HI, CLR_TEXT
    AddCommandButton ws, "set_btn_archive", btnLeft + 160, btnTop + 40, 150, 32, _
        ChrW(128193) & "  ARCHIVE CLOSED", "ArchiveClosed", CLR_PANEL_HI, CLR_TEXT
    AddCommandButton ws, "set_btn_clearlog", btnLeft + 320, btnTop + 40, 150, 32, _
        ChrW(128465) & "  CLEAR LOGS", "ClearLogs", CLR_BAD, vbWhite
    LogInfo "SETTINGS", "Settings sheet built"
End Sub

Private Sub DrawSettingsCard(ws As Worksheet, ByVal title As String, _
        ByVal r As Long, ByVal c As Long, ByVal pairs As Variant)
    Dim hdr As Range
    Set hdr = ws.Range(ws.Cells(r, c), ws.Cells(r, c + 6))
    hdr.Merge
    hdr.Value = title
    hdr.Interior.Color = CLR_PANEL
    hdr.Font.Color = CLR_ACCENT
    hdr.Font.Bold = True
    hdr.Font.Size = 11
    hdr.Font.Name = "Segoe UI Semibold"
    hdr.RowHeight = 24

    Dim i As Long
    For i = LBound(pairs) To UBound(pairs)
        Dim row As Long: row = r + 1 + i
        Dim p As Variant: p = pairs(i)
        ws.Cells(row, c).Value = CStr(p(0))
        ws.Cells(row, c + 1).Value = p(1)
        ws.Range(ws.Cells(row, c), ws.Cells(row, c + 1)).Resize(, 7).Merge
        ws.Cells(row, c).Resize(, 3).Interior.Color = CLR_PANEL_HI
        ws.Cells(row, c).Resize(, 3).Font.Color = CLR_MUTED
        ws.Cells(row, c + 3).Resize(, 4).Interior.Color = CLR_PANEL_HI
        ws.Cells(row, c + 3).Resize(, 4).Font.Color = CLR_TEXT
        ws.Cells(row, c + 3).Resize(, 4).Font.Bold = True
    Next i
End Sub

'==============================================================================
'                          LOGS SHEET
'==============================================================================
Public Sub BuildLogsSheet()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_LOGS)
    ws.Cells.Clear
    ws.Cells.Font.Name = "Consolas"
    ws.Cells.Font.Size = 9
    ws.Cells.Interior.Color = CLR_BG
    ws.Cells.Font.Color = CLR_TEXT

    ws.Range("A1:D1").Value = Array("Timestamp", "Level", "Source", "Message")
    With ws.Range("A1:D1")
        .Interior.Color = CLR_PANEL
        .Font.Color = CLR_ACCENT
        .Font.Bold = True
        .Font.Name = "Segoe UI"
    End With
    ws.Range("A1").ColumnWidth = 22
    ws.Range("B1").ColumnWidth = 8
    ws.Range("C1").ColumnWidth = 14
    ws.Range("D1").ColumnWidth = 100
    ws.Rows("1:1").RowHeight = 22
End Sub

Public Sub ClearLogs()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(SHT_LOGS)
    Dim lr As Long: lr = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    If lr > 1 Then ws.Range("A2:D" & lr).ClearContents
    LogInfo "LOGS", "Logs cleared by user"
End Sub

'==============================================================================
'                          HIDDEN CONFIG
'==============================================================================
Public Sub BuildHiddenConfig()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_HIDDEN)
    ws.Cells.Clear
    ws.Cells(1, 1).Value = "Key": ws.Cells(1, 2).Value = "Value"
    ws.Range("A2").Value = "InitialBuild":  ws.Range("B2").Value = "Dark"          ' theme
    ws.Range("A3").Value = "AdminUser":     ws.Range("B3").Value = "admin"
    ws.Range("A4").Value = "AdminPass":     ws.Range("B4").Value = "admin"
    ws.Range("A5").Value = "MisRecipients": ws.Range("B5").Value = ""
    ws.Range("A6").Value = "CurrentUser":   ws.Range("B6").Value = ""
    ws.Range("A7").Value = "CurrentRole":   ws.Range("B7").Value = ""
    ws.Range("A8").Value = "LastLoginAt":   ws.Range("B8").Value = ""
    ws.Range("A9").Value = "BuildAt":       ws.Range("B9").Value = Now
    ws.Range("A10").Value = "BuildVersion": ws.Range("B10").Value = APP_VERSION
    ws.Visible = xlSheetVeryHidden
End Sub

'==============================================================================
'                          BACKUP & ARCHIVE
'==============================================================================
Public Sub BackupWorkbook()
    On Error GoTo Fail
    Dim path As String, dest As String
    path = ThisWorkbook.Path
    If Len(path) = 0 Then path = Environ$("USERPROFILE") & "\Documents"
    dest = path & Application.PathSeparator & "Backup_" & _
           Replace(ThisWorkbook.name, ".xlsm", "") & "_" & _
           Format(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs dest
    LogInfo "BACKUP", "Saved -> " & dest
    MsgBox "Backup saved:" & vbCrLf & dest, vbInformation, APP_SHORT
    Exit Sub
Fail:
    LogError "BACKUP", Err.Number, Err.Description
    MsgBox "Backup failed: " & Err.Description, vbExclamation, APP_SHORT
End Sub

Public Sub ArchiveClosed()
    On Error GoTo Fail
    Dim arr As Variant: arr = ReadCleanArrayPub
    If IsEmpty(arr) Then Exit Sub
    Dim ws As Worksheet
    Dim archName As String: archName = "Archive_" & Format(Now, "yyyymmdd")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(archName)
    On Error GoTo Fail
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = archName
    End If
    ws.Cells.Clear

    ' Copy header from Clean_Data
    Dim cd As Worksheet: Set cd = ThisWorkbook.Worksheets(SHT_CLEAN)
    cd.Range(cd.Cells(1, 1), cd.Cells(1, CD_TOTAL_COLS)).Copy ws.Range("A1")
    Dim outRow As Long: outRow = 2
    Dim r As Long
    For r = 2 To UBound(arr, 1)
        If CStr(arr(r, CD_NORM_STATUS)) = "Closed" Then
            cd.Rows(r).Copy ws.Rows(outRow)
            outRow = outRow + 1
        End If
    Next r
    ws.Columns.AutoFit
    LogInfo "ARCHIVE", "Archived " & (outRow - 2) & " closed tickets to " & archName
    MsgBox (outRow - 2) & " closed tickets archived to sheet '" & archName & "'", vbInformation, APP_SHORT
    Exit Sub
Fail:
    LogError "ARCHIVE", Err.Number, Err.Description
End Sub
