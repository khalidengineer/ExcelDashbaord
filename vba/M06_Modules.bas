Attribute VB_Name = "M06_Modules"
'==============================================================================
' MODULE      : M06_Modules
' DESCRIPTION : Builds the 8 specialized analytics module sheets that the user
'               navigates between via the side rail.  Each module follows a
'               common "Hero Header > KPI Strip > Charts > Drill Table" layout
'               so the experience feels like Power BI tabs.
'==============================================================================
Option Explicit
Option Compare Text

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
    M04_DashboardUI.ResetSheetForDashboard ws
    M04_DashboardUI.PaintCanvas ws
    M04_DashboardUI.BuildHeaderBar ws, "EXECUTIVE VIEW", _
        "Board-level operational scorecard with health radar"
    M07_Interaction.BuildNavigationRail ws, SHT_EXEC

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
    M05_Charts.BuildOpsRadar ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 380, 260
    M05_Charts.BuildTrendChart ws, _
        ws.Cells(11, 12).Left, ws.Cells(11, 12).Top, 480, 260, "Daily Volume + 3-Day Forecast"
    M05_Charts.BuildSlaGauge ws, _
        ws.Cells(11, 24).Left, ws.Cells(11, 24).Top, 240, 260
    M05_Charts.BuildSeverityDonut ws, _
        ws.Cells(26, 2).Left, ws.Cells(26, 2).Top, 380, 240, "Severity Mix"
    M05_Charts.BuildIssueTypeChart ws, _
        ws.Cells(26, 12).Left, ws.Cells(26, 12).Top, 380, 240
    M05_Charts.BuildLocationVolumeChart ws, _
        ws.Cells(26, 21).Left, ws.Cells(26, 21).Top, 360, 240
End Sub

'==============================================================================
'                          2. SLA INTELLIGENCE
'==============================================================================
Public Sub BuildSlaIntelligence()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_SLA)
    M04_DashboardUI.ResetSheetForDashboard ws
    M04_DashboardUI.PaintCanvas ws
    M04_DashboardUI.BuildHeaderBar ws, "SLA INTELLIGENCE CENTER", _
        "Compliance, breach prediction and recovery insights"
    M07_Interaction.BuildNavigationRail ws, SHT_SLA

    Dim cards As Variant
    cards = Array( _
        Array("SLA %", "SLA_Compliance", "0.0%", CLR_GOOD, ChrW(10003)), _
        Array("Breached", "BreachedTickets", "0", CLR_BAD, ChrW(9888)), _
        Array("At Risk", "AtRiskTickets", "0", CLR_WARN, ChrW(9203)), _
        Array("MTTR (h)", "MTTR_Hours", "0.0", CLR_ACCENT2, ChrW(8987)), _
        Array("Ack (h)", "Ack_Hours", "0.0", CLR_ACCENT, ChrW(9991)), _
        Array("Crit Open", "CriticalOpen", "0", CLR_BAD, ChrW(9760)))
    DrawCardRow ws, "sla", cards, 4, 6

    M05_Charts.BuildSlaGauge ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 280, 260
    M05_Charts.BuildSlaSeverityChart ws, _
        ws.Cells(11, 9).Left, ws.Cells(11, 9).Top, 480, 260
    M05_Charts.BuildAgingFunnel ws, _
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
    M04_DashboardUI.ResetSheetForDashboard ws
    M04_DashboardUI.PaintCanvas ws
    M04_DashboardUI.BuildHeaderBar ws, "INCIDENT WAR ROOM", _
        "Active incidents, escalations and root-cause hotlist"
    M07_Interaction.BuildNavigationRail ws, SHT_INC

    Dim cards As Variant
    cards = Array( _
        Array("Open", "OpenTickets", "0", CLR_WARN, ChrW(9888)), _
        Array("Critical Open", "CriticalOpen", "0", CLR_BAD, ChrW(9760)), _
        Array("Escalation %", "Escalation_Rate", "0.0%", CLR_BAD, ChrW(8607)), _
        Array("Repeat", "RepeatTickets", "0", CLR_PURPLE, ChrW(8634)), _
        Array("Aging > 3d", "AgingOver3d", "0", CLR_WARN, ChrW(9203)), _
        Array("Backlog %", "Backlog_Index", "0.0%", CLR_ACCENT2, ChrW(9776)))
    DrawCardRow ws, "inc", cards, 4, 6

    M05_Charts.BuildSeverityStatusHeat ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 480, 250
    M05_Charts.BuildAgingFunnel ws, _
        ws.Cells(11, 13).Left, ws.Cells(11, 13).Top, 380, 250, "Aging Funnel"
    M05_Charts.BuildHourHeatChart ws, _
        ws.Cells(11, 22).Left, ws.Cells(11, 22).Top, 360, 250

    ' Open Critical drill (live console)
    BuildOpenIncidentsTable ws, 26, 2
End Sub

Private Sub BuildOpenIncidentsTable(ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)
    PaintSectionTitle ws, startRow, startCol, 22, "ACTIVE INCIDENTS  -  Open & Severity Critical/High"

    Dim arr As Variant: arr = M03_KpiEngine.ReadCleanArrayPub
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
            ws.Cells(outRow, startCol + 1).Interior.Color = M05_Charts.SeverityColor(CStr(arr(r, CD_NORM_SEVERITY)))
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
    M04_DashboardUI.ResetSheetForDashboard ws
    M04_DashboardUI.PaintCanvas ws
    M04_DashboardUI.BuildHeaderBar ws, "AGENT PRODUCTIVITY CENTER", _
        "Resolver workload, MTTR and efficiency leaderboard"
    M07_Interaction.BuildNavigationRail ws, SHT_AGENT

    Dim cards As Variant
    cards = Array( _
        Array("Closed", "ClosedTickets", "0", CLR_GOOD, ChrW(10003)), _
        Array("MTTR (h)", "MTTR_Hours", "0.0", CLR_ACCENT2, ChrW(8987)), _
        Array("Eff Score", "Agent_Efficiency", "0", CLR_ACCENT, ChrW(9889)), _
        Array("Auto Close %", "Auto_Closure_Rate", "0.0%", CLR_PURPLE, ChrW(8634)), _
        Array("Util %", "Support_Utilization", "0", CLR_ACCENT2, ChrW(9776)), _
        Array("Closure %", "ClosureRate", "0.0%", CLR_GOOD, ChrW(10004)))
    DrawCardRow ws, "agt", cards, 4, 6

    M05_Charts.BuildResolverChart ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 480, 280
    M05_Charts.BuildResolverEfficiencyChart ws, _
        ws.Cells(11, 13).Left, ws.Cells(11, 13).Top, 480, 280

    ' Resolver leaderboard table
    BuildResolverLeaderboard ws, 26, 2
End Sub

Private Sub BuildResolverLeaderboard(ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)
    PaintSectionTitle ws, startRow, startCol, 22, "RESOLVER LEADERBOARD"
    Dim r1 As Long, r2 As Long
    M05_Charts.LocateBlock "AGG: Resolver MTTR", r1, r2
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
    M04_DashboardUI.ResetSheetForDashboard ws
    M04_DashboardUI.PaintCanvas ws
    M04_DashboardUI.BuildHeaderBar ws, "APPLICATION HEALTH CENTER", _
        "Issue type, location and equipment health monitoring"
    M07_Interaction.BuildNavigationRail ws, SHT_APP

    Dim cards As Variant
    cards = Array( _
        Array("Total", "TotalTickets", "0", CLR_ACCENT, ChrW(8505)), _
        Array("Critical %", "Critical_Failure_Rate", "0.0%", CLR_BAD, ChrW(9760)), _
        Array("Stability", "Stability_Score", "0", CLR_GOOD, ChrW(9776)), _
        Array("Risk", "Risk_Score", "0", CLR_BAD, ChrW(9889)), _
        Array("Downtime Risk", "Downtime_Risk", "0", CLR_WARN, ChrW(9203)), _
        Array("Impact", "Impact_Score", "0.0", CLR_PURPLE, ChrW(9889)))
    DrawCardRow ws, "app", cards, 4, 6

    M05_Charts.BuildIssueTypeChart ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 380, 260
    M05_Charts.BuildLocationVolumeChart ws, _
        ws.Cells(11, 11).Left, ws.Cells(11, 11).Top, 460, 260
    M05_Charts.BuildLocationRiskChart ws, _
        ws.Cells(11, 22).Left, ws.Cells(11, 22).Top, 380, 260
End Sub

'==============================================================================
'                          6. RCA ANALYTICS (Root Cause)
'==============================================================================
Public Sub BuildRcaAnalytics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_RCA)
    M04_DashboardUI.ResetSheetForDashboard ws
    M04_DashboardUI.PaintCanvas ws
    M04_DashboardUI.BuildHeaderBar ws, "RCA ANALYTICS  -  Root Cause Tree", _
        "Auto-clustered themes with recurrence and recommendations"
    M07_Interaction.BuildNavigationRail ws, SHT_RCA

    Dim cards As Variant
    cards = Array( _
        Array("Repeat", "RepeatTickets", "0", CLR_PURPLE, ChrW(8634)), _
        Array("Repeat %", "Repeat_Rate", "0.0%", CLR_PURPLE, ChrW(8635)), _
        Array("Categories", "TotalTickets", "0", CLR_ACCENT, ChrW(9926)), _
        Array("Risk", "Risk_Score", "0", CLR_BAD, ChrW(9889)), _
        Array("Aging > 3d", "AgingOver3d", "0", CLR_WARN, ChrW(9203)), _
        Array("Critical Open", "CriticalOpen", "0", CLR_BAD, ChrW(9760)))
    DrawCardRow ws, "rca", cards, 4, 6

    M05_Charts.BuildCategoryBars ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 480, 280, "Auto-Clustered Categories"
    BuildRcaTree ws, 11, 13

    ' Recommendations panel from KPI sheet
    BuildAiRecommendationsTable ws, 26, 2
End Sub

Private Sub BuildRcaTree(ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long)
    ' Visual root-cause tree using shapes - parent "All Tickets" -> child categories
    Dim arr As Variant: arr = M03_KpiEngine.ReadCleanArrayPub
    If IsEmpty(arr) Then Exit Sub
    Dim cat As Object: Set cat = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To UBound(arr, 1)
        k = CStr(arr(r, CD_CATEGORY))
        If cat.Exists(k) Then cat(k) = cat(k) + 1 Else cat(k) = 1
    Next r
    Dim keys As Variant: keys = cat.keys
    M02_DataEngine.SortDictDesc cat, keys

    PaintSectionTitle ws, startRow, startCol, 16, "ROOT CAUSE TREE  (auto-clustered)"

    Dim baseLeft As Double, baseTop As Double
    baseLeft = ws.Cells(startRow + 1, startCol).Left
    baseTop = ws.Cells(startRow + 1, startCol).Top

    ' Root node
    Dim root As Shape
    Set root = M04_DashboardUI.AddShape(ws, msoShapeRoundedRectangle, "rca_root", _
        baseLeft + 130, baseTop + 8, 140, 36)
    M04_DashboardUI.StyleShape root, CLR_ACCENT, 0
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
        Set child = M04_DashboardUI.AddShape(ws, msoShapeRoundedRectangle, "rca_c" & n, _
            baseLeft + 60, childY, 280, childH)
        M04_DashboardUI.StyleShape child, CLR_PANEL_HI, 0
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
    M04_DashboardUI.ResetSheetForDashboard ws
    M04_DashboardUI.PaintCanvas ws
    M04_DashboardUI.BuildHeaderBar ws, "OPERATIONAL FORECAST CENTER", _
        "3-day projected demand, peak hours, and capacity planning"
    M07_Interaction.BuildNavigationRail ws, SHT_FORECAST

    Dim cards As Variant
    Dim peak As Variant: peak = M03_KpiEngine.GetKpi("Peak_Hour")
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
    Set pill = M04_DashboardUI.AddShape(ws, msoShapeRoundedRectangle, "fc_peak", _
        ws.Cells(8, 2).Left, ws.Cells(8, 2).Top, 380, 32)
    M04_DashboardUI.StyleShape pill, CLR_PANEL_HI, 0
    pill.TextFrame2.TextRange.Text = "PEAK INTAKE WINDOW:  " & CStr(peak)
    With pill.TextFrame2.TextRange.Font
        .Bold = msoTrue: .Size = 10: .Name = "Segoe UI"
        .Fill.ForeColor.RGB = CLR_ACCENT
    End With
    pill.TextFrame2.HorizontalAnchor = msoAnchorCenter
    pill.TextFrame2.VerticalAnchor = msoAnchorMiddle

    M05_Charts.BuildForecastChart ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 540, 280
    M05_Charts.BuildHourHeatChart ws, _
        ws.Cells(11, 14).Left, ws.Cells(11, 14).Top, 460, 280
    M05_Charts.BuildDowChart ws, _
        ws.Cells(11, 24).Left, ws.Cells(11, 24).Top, 380, 280

    ' Hour x DOW heatmap (cell-based)
    PaintSectionTitle ws, 26, 2, 28, "HOUR x DAY-OF-WEEK INTAKE HEATMAP"
    M05_Charts.BuildHourDowHeatmap ws, 27, 2
End Sub

'==============================================================================
'                          8. RISK ANALYTICS
'==============================================================================
Public Sub BuildRiskAnalytics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_RISK)
    M04_DashboardUI.ResetSheetForDashboard ws
    M04_DashboardUI.PaintCanvas ws
    M04_DashboardUI.BuildHeaderBar ws, "RISK MONITORING CENTER", _
        "Composite risk, downtime, and high-impact site watch-list"
    M07_Interaction.BuildNavigationRail ws, SHT_RISK

    Dim cards As Variant
    cards = Array( _
        Array("Risk Score", "Risk_Score", "0", CLR_BAD, ChrW(9889)), _
        Array("Downtime", "Downtime_Risk", "0", CLR_BAD, ChrW(9203)), _
        Array("Aging Risk", "Aging_Risk", "0", CLR_WARN, ChrW(9203)), _
        Array("Critical %", "Critical_Failure_Rate", "0.0%", CLR_BAD, ChrW(9760)), _
        Array("Priority Heat", "Priority_Heat", "0.0", CLR_WARN, ChrW(9889)), _
        Array("Stability", "Stability_Score", "0", CLR_GOOD, ChrW(9776)))
    DrawCardRow ws, "rsk", cards, 4, 6

    M05_Charts.BuildLocationRiskChart ws, _
        ws.Cells(11, 2).Left, ws.Cells(11, 2).Top, 460, 280
    M05_Charts.BuildResolverEfficiencyChart ws, _
        ws.Cells(11, 13).Left, ws.Cells(11, 13).Top, 460, 280
    M05_Charts.BuildOpsRadar ws, _
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

    Dim arr As Variant: arr = M03_KpiEngine.ReadCleanArrayPub
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
        ws.Cells(outRow, startCol).Interior.Color = M05_Charts.SeverityColor(CStr(sevs(i)))
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
        M04_DashboardUI.DrawKpiCard ws, _
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

Private Function NzNum(ByVal v As Variant) As Double
    If IsNumeric(v) Then NzNum = CDbl(v) Else NzNum = 0
End Function
