Attribute VB_Name = "M03_KpiEngine"
'==============================================================================
' MODULE      : M03_KpiEngine
' DESCRIPTION : Computes enterprise KPIs and runs AI-style analytics.
'               Outputs land on KPI_Engine sheet as a key/value store that the
'               UI layer reads directly via GetKpi(name).
'==============================================================================
Option Explicit
Option Compare Text

' KPI_Engine layout
'   Col A: Key (string)            Col B: Value (variant)
'   Col D: Section header          Col E..: AI tables (alerts, recos)

'==============================================================================
'                          PUBLIC GETTER
'==============================================================================
Public Function GetKpi(ByVal key As String) As Variant
    Dim ws As Worksheet, lr As Long, i As Long
    Set ws = ThisWorkbook.Worksheets(SHT_KPI)
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    For i = 2 To lr
        If StrComp(CStr(ws.Cells(i, 1).Value), key, vbTextCompare) = 0 Then
            GetKpi = ws.Cells(i, 2).Value
            Exit Function
        End If
    Next i
    GetKpi = ""
End Function

Public Function FmtKpi(ByVal key As String, ByVal numFmt As String) As String
    Dim v As Variant: v = GetKpi(key)
    If IsNumeric(v) Then
        FmtKpi = Format(v, numFmt)
    Else
        FmtKpi = CStr(v)
    End If
End Function

'==============================================================================
'                          KPI COMPUTATION
'==============================================================================
Public Sub ComputeAllKpis()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_KPI)
    ws.Cells.Clear
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 10

    ws.Range("A1:B1").Value = Array("KPI Key", "Value")
    With ws.Range("A1:B1")
        .Interior.Color = CLR_PANEL
        .Font.Color = CLR_ACCENT
        .Font.Bold = True
    End With

    Dim arr As Variant
    arr = ReadCleanArrayPub()
    If IsEmpty(arr) Then Exit Sub
    Dim n As Long: n = UBound(arr, 1) - 1   ' data rows excluding header

    Dim total As Long, opn As Long, closed As Long, breached As Long, atRisk As Long
    Dim crit As Long, criticalOpen As Long, repeatCnt As Long
    Dim sumRes As Double, cntRes As Long, sumImpact As Double, sumRisk As Double
    Dim sevCounts As Object: Set sevCounts = CreateObject("Scripting.Dictionary")
    Dim agingDays3 As Long, agingDays7 As Long, autoClose As Long
    Dim r As Long
    total = n
    For r = 2 To UBound(arr, 1)
        Select Case CStr(arr(r, CD_NORM_STATUS))
            Case "Open":   opn = opn + 1
            Case "Closed": closed = closed + 1
        End Select
        If CStr(arr(r, CD_NORM_SEVERITY)) = "Critical" Then crit = crit + 1
        If CStr(arr(r, CD_NORM_SEVERITY)) = "Critical" And CStr(arr(r, CD_NORM_STATUS)) = "Open" Then criticalOpen = criticalOpen + 1
        Select Case CStr(arr(r, CD_SLA_STATUS))
            Case "Breached": breached = breached + 1
            Case "At Risk":  atRisk = atRisk + 1
        End Select
        If IsNumeric(arr(r, CD_RES_HOURS)) Then
            sumRes = sumRes + CDbl(arr(r, CD_RES_HOURS))
            cntRes = cntRes + 1
        End If
        sumImpact = sumImpact + NzNum(arr(r, CD_IMPACT_SCORE))
        sumRisk = sumRisk + NzNum(arr(r, CD_RISK_SCORE))
        If CStr(arr(r, CD_REPEAT_FLAG)) = "Repeat" Then repeatCnt = repeatCnt + 1
        If IsNumeric(arr(r, CD_AGE_HOURS)) Then
            If CDbl(arr(r, CD_AGE_HOURS)) > 72 Then agingDays3 = agingDays3 + 1
            If CDbl(arr(r, CD_AGE_HOURS)) > 168 Then agingDays7 = agingDays7 + 1
        End If
        ' Self-resolved (resolver = reporter -> auto closure proxy)
        If Len(Trim$(CStr(arr(r, CD_RESOLVER)))) > 0 Then
            If CStr(arr(r, CD_RESOLVER)) = CStr(arr(r, CD_REPORTER)) Then autoClose = autoClose + 1
        End If
        Dim sv As String: sv = CStr(arr(r, CD_NORM_SEVERITY))
        If sevCounts.Exists(sv) Then sevCounts(sv) = sevCounts(sv) + 1 Else sevCounts(sv) = 1
    Next r

    Dim closureRate As Double, slaCompliance As Double, mttr As Double
    closureRate = SafeDiv(closed, total)
    Dim slaTotal As Long, slaMet As Long
    For r = 2 To UBound(arr, 1)
        Select Case CStr(arr(r, CD_SLA_STATUS))
            Case "Met", "On Track": slaMet = slaMet + 1: slaTotal = slaTotal + 1
            Case "Breached", "At Risk": slaTotal = slaTotal + 1
        End Select
    Next r
    slaCompliance = SafeDiv(slaMet, slaTotal)
    mttr = SafeDiv(sumRes, cntRes)

    Dim escalationRate As Double, repeatRate As Double, backlog As Double
    escalationRate = SafeDiv(criticalOpen, opn)
    repeatRate = SafeDiv(repeatCnt, total)
    backlog = SafeDiv(opn, total)

    Dim stability As Double, riskScore As Double, opsHealth As Double
    stability = WorksheetFunction.Max(0, 100 - (criticalOpen * 8) - (atRisk * 4) - (breached * 5))
    If stability > 100 Then stability = 100
    riskScore = WorksheetFunction.Min(100, SafeDiv(sumRisk, total))
    opsHealth = WorksheetFunction.Max(0, 100 - riskScore * 0.6 - (1 - slaCompliance) * 40)

    Dim impactScore As Double, supportUtil As Double, agentEff As Double
    impactScore = SafeDiv(sumImpact, total)
    supportUtil = WorksheetFunction.Min(100, closed * 1.5)
    agentEff = slaCompliance * 60 + closureRate * 40

    Dim cust As Double, criticalFailureRate As Double, agingRisk As Double
    cust = WorksheetFunction.Max(0, 100 - (1 - slaCompliance) * 50 - repeatRate * 100 * 0.4)
    criticalFailureRate = SafeDiv(crit, total)
    agingRisk = WorksheetFunction.Min(100, agingDays3 * 6 + agingDays7 * 10)

    Dim priorityHeat As Double, trendVelocity As Double, ackTime As Double
    priorityHeat = SafeDiv(crit + atRisk * 0.6 + breached * 0.8, total) * 100
    trendVelocity = WeekOverWeek(arr)
    ackTime = mttr * 0.25     ' proxy: avg acknowledgement modelled as 1/4 of MTTR

    Dim downtimeRisk As Double
    downtimeRisk = WorksheetFunction.Min(100, criticalOpen * 12 + breached * 6 + atRisk * 3)

    ' --- Write block ---
    Dim row As Long: row = 2
    Set sevCounts = sevCounts
    PutKpi ws, row, "TotalTickets", total
    PutKpi ws, row, "OpenTickets", opn
    PutKpi ws, row, "ClosedTickets", closed
    PutKpi ws, row, "CriticalTickets", crit
    PutKpi ws, row, "CriticalOpen", criticalOpen
    PutKpi ws, row, "BreachedTickets", breached
    PutKpi ws, row, "AtRiskTickets", atRisk
    PutKpi ws, row, "RepeatTickets", repeatCnt
    PutKpi ws, row, "AgingOver3d", agingDays3
    PutKpi ws, row, "AgingOver7d", agingDays7
    PutKpi ws, row, "AutoClosed", autoClose

    PutKpi ws, row, "ClosureRate", closureRate
    PutKpi ws, row, "SLA_Compliance", slaCompliance
    PutKpi ws, row, "MTTR_Hours", mttr
    PutKpi ws, row, "Ack_Hours", ackTime
    PutKpi ws, row, "Backlog_Index", backlog
    PutKpi ws, row, "Escalation_Rate", escalationRate
    PutKpi ws, row, "Repeat_Rate", repeatRate
    PutKpi ws, row, "Stability_Score", stability
    PutKpi ws, row, "Risk_Score", riskScore
    PutKpi ws, row, "OpsHealth_Score", opsHealth
    PutKpi ws, row, "Impact_Score", impactScore
    PutKpi ws, row, "Downtime_Risk", downtimeRisk
    PutKpi ws, row, "Critical_Failure_Rate", criticalFailureRate
    PutKpi ws, row, "Support_Utilization", supportUtil
    PutKpi ws, row, "Agent_Efficiency", agentEff
    PutKpi ws, row, "Auto_Closure_Rate", SafeDiv(autoClose, closed)
    PutKpi ws, row, "Customer_Satisfaction", cust
    PutKpi ws, row, "Aging_Risk", agingRisk
    PutKpi ws, row, "Priority_Heat", priorityHeat
    PutKpi ws, row, "Trend_Velocity", trendVelocity

    PutKpi ws, row, "Sev_Critical", IIf(sevCounts.Exists("Critical"), sevCounts("Critical"), 0)
    PutKpi ws, row, "Sev_High", IIf(sevCounts.Exists("High"), sevCounts("High"), 0)
    PutKpi ws, row, "Sev_Medium", IIf(sevCounts.Exists("Medium"), sevCounts("Medium"), 0)
    PutKpi ws, row, "Sev_Low", IIf(sevCounts.Exists("Low"), sevCounts("Low"), 0)
    PutKpi ws, row, "Sev_Unspecified", IIf(sevCounts.Exists("Unspecified"), sevCounts("Unspecified"), 0)

    PutKpi ws, row, "LastRefreshed", Now
    ws.Cells(row - 1, 2).NumberFormat = "yyyy-mm-dd hh:mm:ss"

    ws.Columns("A:B").AutoFit
    LogInfo "KPI", "Computed " & row & " KPIs"
End Sub

Private Sub PutKpi(ws As Worksheet, ByRef row As Long, ByVal key As String, ByVal value As Variant)
    ws.Cells(row, 1).Value = key
    ws.Cells(row, 2).Value = value
    ws.Cells(row, 1).Font.Color = CLR_MUTED
    ws.Cells(row, 2).Font.Color = CLR_TEXT
    row = row + 1
End Sub

Private Function SafeDiv(num As Double, den As Double) As Double
    If den = 0 Then SafeDiv = 0 Else SafeDiv = num / den
End Function

Private Function WeekOverWeek(arr As Variant) As Double
    Dim curWeek As Long, prevWeek As Long, r As Long
    Dim today As Date, weekAgo As Date, twoWeekAgo As Date
    today = Date
    weekAgo = today - 7
    twoWeekAgo = today - 14
    For r = 2 To UBound(arr, 1)
        If IsDate(arr(r, CD_REPORT_DT)) Then
            Dim d As Date: d = CDate(arr(r, CD_REPORT_DT))
            If d > weekAgo Then curWeek = curWeek + 1 _
            ElseIf d > twoWeekAgo Then prevWeek = prevWeek + 1
        End If
    Next r
    If prevWeek = 0 Then
        WeekOverWeek = IIf(curWeek > 0, 100, 0)
    Else
        WeekOverWeek = ((curWeek - prevWeek) / prevWeek) * 100
    End If
End Function

Public Function ReadCleanArrayPub() As Variant
    Dim ws As Worksheet, lr As Long
    Set ws = ThisWorkbook.Worksheets(SHT_CLEAN)
    lr = ws.Cells(ws.Rows.Count, CD_ISSUE_ID).End(xlUp).Row
    If lr < 2 Then ReadCleanArrayPub = Empty: Exit Function
    ReadCleanArrayPub = ws.Range(ws.Cells(1, 1), ws.Cells(lr, CD_TOTAL_COLS)).Value
End Function

'==============================================================================
'                          AI-STYLE ANALYTICS
'==============================================================================
Public Sub RunAiAnalytics()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_KPI)
    Dim arr As Variant: arr = ReadCleanArrayPub()
    If IsEmpty(arr) Then Exit Sub

    ' --- Smart Alerts (rule-based AI) ---
    Dim alerts As Collection: Set alerts = New Collection
    GenerateAlerts arr, alerts

    ' --- AI Recommendations ---
    Dim recos As Collection: Set recos = New Collection
    GenerateRecommendations arr, recos

    ' --- Predicted breaches (open tickets approaching SLA) ---
    Dim preds As Collection: Set preds = New Collection
    PredictBreaches arr, preds

    ' --- High-risk applications/locations ---
    Dim hra As Collection: Set hra = New Collection
    DetectHighRiskLocations arr, hra

    ' --- Peak hour detection ---
    Dim peak As String
    peak = DetectPeakHour(arr)

    ' --- Smart severity suggestions for unspecified ---
    Dim suggs As Collection: Set suggs = New Collection
    SuggestSeverities arr, suggs

    ' Write AI region starting at column D
    Dim cBase As Long: cBase = 4
    Dim row As Long: row = 1
    ws.Cells(row, cBase).Value = "AI ALERTS"
    StyleAiHeader ws.Cells(row, cBase), 4
    row = row + 1
    ws.Cells(row, cBase).Resize(1, 4).Value = Array("Severity", "Module", "Title", "Recommendation")
    StyleAiSubheader ws.Range(ws.Cells(row, cBase), ws.Cells(row, cBase + 3))
    Dim it As Variant, i As Long
    i = 0
    For Each it In alerts
        row = row + 1: i = i + 1
        ws.Cells(row, cBase).Value = it(0)
        ws.Cells(row, cBase + 1).Value = it(1)
        ws.Cells(row, cBase + 2).Value = it(2)
        ws.Cells(row, cBase + 3).Value = it(3)
        ColorAlertRow ws.Range(ws.Cells(row, cBase), ws.Cells(row, cBase + 3)), CStr(it(0))
    Next it

    row = row + 2
    ws.Cells(row, cBase).Value = "AI RECOMMENDATIONS"
    StyleAiHeader ws.Cells(row, cBase), 4
    row = row + 1
    ws.Cells(row, cBase).Resize(1, 3).Value = Array("Theme", "Insight", "Action")
    StyleAiSubheader ws.Range(ws.Cells(row, cBase), ws.Cells(row, cBase + 2))
    For Each it In recos
        row = row + 1
        ws.Cells(row, cBase).Value = it(0)
        ws.Cells(row, cBase + 1).Value = it(1)
        ws.Cells(row, cBase + 2).Value = it(2)
    Next it

    row = row + 2
    ws.Cells(row, cBase).Value = "PREDICTED SLA BREACHES"
    StyleAiHeader ws.Cells(row, cBase), 4
    row = row + 1
    ws.Cells(row, cBase).Resize(1, 4).Value = Array("Issue ID", "Title", "Age (h)", "SLA (h)")
    StyleAiSubheader ws.Range(ws.Cells(row, cBase), ws.Cells(row, cBase + 3))
    For Each it In preds
        row = row + 1
        ws.Cells(row, cBase).Value = it(0)
        ws.Cells(row, cBase + 1).Value = it(1)
        ws.Cells(row, cBase + 2).Value = it(2)
        ws.Cells(row, cBase + 3).Value = it(3)
        ws.Cells(row, cBase + 2).NumberFormat = "0.0"
    Next it

    row = row + 2
    ws.Cells(row, cBase).Value = "HIGH-RISK LOCATIONS"
    StyleAiHeader ws.Cells(row, cBase), 4
    row = row + 1
    ws.Cells(row, cBase).Resize(1, 3).Value = Array("Location", "Tickets", "Avg Risk")
    StyleAiSubheader ws.Range(ws.Cells(row, cBase), ws.Cells(row, cBase + 2))
    For Each it In hra
        row = row + 1
        ws.Cells(row, cBase).Value = it(0)
        ws.Cells(row, cBase + 1).Value = it(1)
        ws.Cells(row, cBase + 2).Value = it(2)
        ws.Cells(row, cBase + 2).NumberFormat = "0.0"
    Next it

    row = row + 2
    ws.Cells(row, cBase).Value = "SUGGESTED SEVERITY (for blanks)"
    StyleAiHeader ws.Cells(row, cBase), 4
    row = row + 1
    ws.Cells(row, cBase).Resize(1, 3).Value = Array("Issue ID", "Title", "Suggested")
    StyleAiSubheader ws.Range(ws.Cells(row, cBase), ws.Cells(row, cBase + 2))
    For Each it In suggs
        row = row + 1
        ws.Cells(row, cBase).Value = it(0)
        ws.Cells(row, cBase + 1).Value = it(1)
        ws.Cells(row, cBase + 2).Value = it(2)
    Next it

    ' Peak hour as KPI
    Dim lr As Long
    lr = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    ws.Cells(lr, 1).Value = "Peak_Hour"
    ws.Cells(lr, 2).Value = peak

    ws.Columns(cBase).ColumnWidth = 14
    ws.Columns(cBase + 1).ColumnWidth = 28
    ws.Columns(cBase + 2).ColumnWidth = 50
    ws.Columns(cBase + 3).ColumnWidth = 38
    LogInfo "AI", "Generated " & alerts.Count & " alerts, " & recos.Count & " recos, " & preds.Count & " breach predictions"
End Sub

'------------------------------------------------------------------------------
'                          ALERT GENERATOR (rule-based)
'------------------------------------------------------------------------------
Private Sub GenerateAlerts(arr As Variant, alerts As Collection)
    Dim total As Long: total = UBound(arr, 1) - 1
    Dim opn As Long, crit As Long, criticalOpen As Long, breached As Long, atRisk As Long
    Dim repeatCnt As Long, agingDays3 As Long, r As Long
    For r = 2 To UBound(arr, 1)
        If CStr(arr(r, CD_NORM_STATUS)) = "Open" Then opn = opn + 1
        If CStr(arr(r, CD_NORM_SEVERITY)) = "Critical" Then crit = crit + 1
        If CStr(arr(r, CD_NORM_SEVERITY)) = "Critical" And CStr(arr(r, CD_NORM_STATUS)) = "Open" Then criticalOpen = criticalOpen + 1
        If CStr(arr(r, CD_SLA_STATUS)) = "Breached" Then breached = breached + 1
        If CStr(arr(r, CD_SLA_STATUS)) = "At Risk" Then atRisk = atRisk + 1
        If CStr(arr(r, CD_REPEAT_FLAG)) = "Repeat" Then repeatCnt = repeatCnt + 1
        If IsNumeric(arr(r, CD_AGE_HOURS)) Then
            If CDbl(arr(r, CD_AGE_HOURS)) > 72 And CStr(arr(r, CD_NORM_STATUS)) = "Open" Then agingDays3 = agingDays3 + 1
        End If
    Next r

    If criticalOpen > 0 Then
        alerts.Add Array("Critical", "Incident War Room", _
            criticalOpen & " Critical tickets still OPEN", _
            "Convene war-room standup immediately. Assign senior resolver and post hourly status.")
    End If
    If breached >= 1 Then
        alerts.Add Array("High", "SLA Intelligence", _
            breached & " ticket(s) have breached SLA", _
            "Escalate to delivery lead. Trigger root-cause review on next business day.")
    End If
    If atRisk >= 3 Then
        alerts.Add Array("Medium", "SLA Intelligence", _
            atRisk & " tickets approaching SLA breach", _
            "Pre-emptive reallocation of resolver capacity. Notify affected stakeholders.")
    End If
    If agingDays3 >= 2 Then
        alerts.Add Array("Medium", "Ticket Aging Center", _
            agingDays3 & " open tickets older than 72 hours", _
            "Run aging triage. Auto-close stale tickets after stakeholder confirmation.")
    End If
    If repeatCnt > 0 Then
        alerts.Add Array("Medium", "Repeat Incident", _
            repeatCnt & " repeat incidents detected", _
            "Open Problem records and schedule preventive maintenance to break recurrence.")
    End If
    Dim sla As Variant: sla = GetKpi("SLA_Compliance")
    If IsNumeric(sla) Then
        If CDbl(sla) < 0.85 Then
            alerts.Add Array("High", "Executive Command", _
                "SLA Compliance at " & Format(CDbl(sla), "0.0%"), _
                "Target is 95%+. Run SLA recovery sprint and review resolver workload distribution.")
        End If
    End If
End Sub

'------------------------------------------------------------------------------
'                       RECOMMENDATIONS GENERATOR
'------------------------------------------------------------------------------
Private Sub GenerateRecommendations(arr As Variant, recos As Collection)
    ' Theme 1: Top failing category
    Dim cat As Object: Set cat = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To UBound(arr, 1)
        k = CStr(arr(r, CD_CATEGORY))
        If cat.Exists(k) Then cat(k) = cat(k) + 1 Else cat(k) = 1
    Next r
    Dim ks As Variant: ks = cat.keys
    SortDictDesc cat, ks
    If UBound(ks) >= 0 Then
        recos.Add Array("Top Failure Theme", _
            CStr(ks(0)) & " accounts for " & cat(ks(0)) & " tickets", _
            "Initiate a quarterly preventive-maintenance contract for this category.")
    End If

    ' Theme 2: Repeat reporters
    Dim rep As Object: Set rep = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(arr, 1)
        k = CStr(arr(r, CD_REPORTER))
        If rep.Exists(k) Then rep(k) = rep(k) + 1 Else rep(k) = 1
    Next r
    ks = rep.keys: SortDictDesc rep, ks
    If UBound(ks) >= 0 And rep(ks(0)) >= 3 Then
        recos.Add Array("Reporter Concentration", _
            CStr(ks(0)) & " has raised " & rep(ks(0)) & " tickets", _
            "Run a site walk-through with this reporter to identify systemic issues at their location.")
    End If

    ' Theme 3: Resolver workload imbalance
    Dim res As Object: Set res = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(arr, 1)
        k = Trim$(CStr(arr(r, CD_RESOLVER)))
        If Len(k) > 0 Then
            If res.Exists(k) Then res(k) = res(k) + 1 Else res(k) = 1
        End If
    Next r
    ks = res.keys: SortDictDesc res, ks
    If UBound(ks) >= 1 Then
        Dim top1 As Long, top2 As Long
        top1 = res(ks(0)): top2 = res(ks(1))
        If top1 >= 2 * top2 And top1 >= 4 Then
            recos.Add Array("Resolver Load", _
                CStr(ks(0)) & " carries " & top1 & " tickets vs next agent " & top2, _
                "Rebalance ticket routing rules and consider hiring/cross-training.")
        End If
    End If

    ' Theme 4: Off-hours volume
    Dim biz As Long, off As Long
    For r = 2 To UBound(arr, 1)
        If CStr(arr(r, CD_REPORT_BIZ)) = "Business" Then biz = biz + 1
        If CStr(arr(r, CD_REPORT_BIZ)) = "Off-Hours" Then off = off + 1
    Next r
    If (biz + off) > 0 Then
        Dim offRate As Double: offRate = off / (biz + off)
        If offRate > 0.4 Then
            recos.Add Array("Off-Hours Demand", _
                Format(offRate, "0.0%") & " of tickets raised off-hours", _
                "Stand up a 24x7 on-call rota with priority-based paging.")
        End If
    End If

    ' Theme 5: Severity hygiene
    Dim unspec As Long
    For r = 2 To UBound(arr, 1)
        If CStr(arr(r, CD_NORM_SEVERITY)) = "Unspecified" Then unspec = unspec + 1
    Next r
    If unspec > 0 Then
        recos.Add Array("Data Hygiene", _
            unspec & " tickets have no severity tagged", _
            "Make Severity a mandatory intake field. Backfill using the Smart Severity Classifier.")
    End If

    ' Theme 6: Critical concentration
    Dim sevCrit As Object: Set sevCrit = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(arr, 1)
        If CStr(arr(r, CD_NORM_SEVERITY)) = "Critical" Then
            k = CStr(arr(r, CD_REP_LOC))
            If sevCrit.Exists(k) Then sevCrit(k) = sevCrit(k) + 1 Else sevCrit(k) = 1
        End If
    Next r
    ks = sevCrit.keys: SortDictDesc sevCrit, ks
    If UBound(ks) >= 0 And sevCrit(ks(0)) >= 3 Then
        recos.Add Array("Critical Hot-Spot", _
            CStr(ks(0)) & " has " & sevCrit(ks(0)) & " critical tickets", _
            "Schedule on-site senior engineer audit. Add this site to executive watch-list.")
    End If
End Sub

'------------------------------------------------------------------------------
'                       BREACH PREDICTOR (top open at-risk)
'------------------------------------------------------------------------------
Private Sub PredictBreaches(arr As Variant, preds As Collection)
    Dim r As Long, age As Double, sla As Double, score As Double
    Dim cand As Collection: Set cand = New Collection
    Dim scores As Collection: Set scores = New Collection
    For r = 2 To UBound(arr, 1)
        If CStr(arr(r, CD_NORM_STATUS)) = "Open" Then
            age = NzNum(arr(r, CD_AGE_HOURS))
            sla = NzNum(arr(r, CD_SLA_TARGET))
            If sla > 0 Then
                score = age / sla
                cand.Add Array(arr(r, CD_ISSUE_ID), arr(r, CD_ISSUE_TITLE), age, sla)
                scores.Add score
            End If
        End If
    Next r
    ' simple selection sort top 10
    Dim i As Long, j As Long, tmp As Variant, tmpS As Double
    Dim arrCand() As Variant, arrScores() As Double, n As Long
    n = cand.Count
    If n = 0 Then Exit Sub
    ReDim arrCand(1 To n): ReDim arrScores(1 To n)
    For i = 1 To n
        arrCand(i) = cand(i): arrScores(i) = scores(i)
    Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If arrScores(j) > arrScores(i) Then
                tmpS = arrScores(i): arrScores(i) = arrScores(j): arrScores(j) = tmpS
                tmp = arrCand(i): arrCand(i) = arrCand(j): arrCand(j) = tmp
            End If
        Next j
    Next i
    Dim top As Long: top = WorksheetFunction.Min(10, n)
    For i = 1 To top
        preds.Add arrCand(i)
    Next i
End Sub

'------------------------------------------------------------------------------
'                  HIGH-RISK LOCATION DETECTOR
'------------------------------------------------------------------------------
Private Sub DetectHighRiskLocations(arr As Variant, hra As Collection)
    Dim cnt As Object, sumR As Object
    Set cnt = CreateObject("Scripting.Dictionary")
    Set sumR = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To UBound(arr, 1)
        k = NzString(arr(r, CD_REP_LOC), "(blank)")
        If Not cnt.Exists(k) Then cnt(k) = 0: sumR(k) = 0
        cnt(k) = cnt(k) + 1
        sumR(k) = sumR(k) + NzNum(arr(r, CD_RISK_SCORE))
    Next r
    Dim ks As Variant: ks = cnt.keys
    Dim avgR As Object: Set avgR = CreateObject("Scripting.Dictionary")
    Dim i As Long
    For i = 0 To UBound(ks)
        avgR(ks(i)) = sumR(ks(i)) / cnt(ks(i))
    Next i
    SortDictDesc avgR, ks
    Dim top As Long: top = WorksheetFunction.Min(10, UBound(ks))
    For i = 0 To top
        hra.Add Array(ks(i), cnt(ks(i)), avgR(ks(i)))
    Next i
End Sub

'------------------------------------------------------------------------------
'                       PEAK HOUR DETECTOR
'------------------------------------------------------------------------------
Private Function DetectPeakHour(arr As Variant) As String
    Dim h(0 To 23) As Long, r As Long
    For r = 2 To UBound(arr, 1)
        If IsNumeric(arr(r, CD_HOUR)) Then h(CLng(arr(r, CD_HOUR))) = h(CLng(arr(r, CD_HOUR))) + 1
    Next r
    Dim maxIdx As Long, i As Long
    For i = 0 To 23
        If h(i) > h(maxIdx) Then maxIdx = i
    Next i
    DetectPeakHour = Format(maxIdx, "00") & ":00 - " & Format(maxIdx + 1, "00") & ":00 (" & h(maxIdx) & " tickets)"
End Function

'------------------------------------------------------------------------------
'                       SMART SEVERITY CLASSIFIER (for blanks)
'------------------------------------------------------------------------------
Private Sub SuggestSeverities(arr As Variant, suggs As Collection)
    Dim r As Long, t As String, sug As String, count As Long
    For r = 2 To UBound(arr, 1)
        If CStr(arr(r, CD_NORM_SEVERITY)) = "Unspecified" Then
            t = LCase$(CStr(arr(r, CD_ISSUE_TITLE)) & " " & CStr(arr(r, CD_ISSUE_DESC)))
            sug = SuggestSeverityFromText(t)
            suggs.Add Array(arr(r, CD_ISSUE_ID), arr(r, CD_ISSUE_TITLE), sug)
            count = count + 1
            If count >= 12 Then Exit For
        End If
    Next r
End Sub

Public Function SuggestSeverityFromText(ByVal t As String) As String
    Dim s As String: s = LCase$(t)
    If InStr(s, "fire") > 0 Or InStr(s, "burn") > 0 Or InStr(s, "smoke") > 0 Or InStr(s, "flood") > 0 Or InStr(s, "leakage") > 0 Then SuggestSeverityFromText = "Critical": Exit Function
    If InStr(s, "not working") > 0 Or InStr(s, "stopped") > 0 Or InStr(s, "down") > 0 Or InStr(s, "fail") > 0 Or InStr(s, "broken") > 0 Then SuggestSeverityFromText = "High": Exit Function
    If InStr(s, "issue") > 0 Or InStr(s, "problem") > 0 Or InStr(s, "noise") > 0 Or InStr(s, "leak") > 0 Then SuggestSeverityFromText = "Medium": Exit Function
    SuggestSeverityFromText = "Low"
End Function

'==============================================================================
'                          STYLE HELPERS
'==============================================================================
Private Sub StyleAiHeader(c As Range, ByVal span As Long)
    With c.Resize(1, span)
        .Merge
        .Value = c.Value
        .Interior.Color = CLR_PANEL
        .Font.Color = CLR_ACCENT
        .Font.Bold = True
        .Font.Size = 12
        .RowHeight = 26
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub StyleAiSubheader(rng As Range)
    With rng
        .Interior.Color = CLR_BORDER
        .Font.Color = CLR_TEXT
        .Font.Bold = True
        .Font.Size = 9
    End With
End Sub

Private Sub ColorAlertRow(rng As Range, ByVal sev As String)
    Dim col As Long
    Select Case sev
        Case "Critical": col = CLR_BAD
        Case "High":     col = CLR_WARN
        Case "Medium":   col = CLR_ACCENT2
        Case Else:       col = CLR_PANEL_HI
    End Select
    rng.Cells(1, 1).Interior.Color = col
    rng.Cells(1, 1).Font.Color = vbWhite
    rng.Cells(1, 1).Font.Bold = True
    rng.Resize(, rng.Columns.Count).Font.Size = 9
End Sub
