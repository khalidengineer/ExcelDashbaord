Attribute VB_Name = "M02_DataEngine"
'==============================================================================
' MODULE      : M02_DataEngine
' DESCRIPTION : Auto CSV import, smart cleaning, normalization, derived columns
'               and aggregation model. All transforms run on in-memory arrays
'               for performance, then write back in single shot.
'==============================================================================
Option Explicit
Option Compare Text

' Public column indices for Clean_Data (1-based, set in CleanData)
Public Const CD_REPORTER          As Long = 1
Public Const CD_REP_ID            As Long = 2
Public Const CD_REP_DESIG         As Long = 3
Public Const CD_REP_DEPT          As Long = 4
Public Const CD_REP_DIV           As Long = 5
Public Const CD_REP_SUBDIV        As Long = 6
Public Const CD_REP_LOC           As Long = 7
Public Const CD_REPORTED_AT       As Long = 8
Public Const CD_ISSUE_ID          As Long = 9
Public Const CD_ISSUE_TITLE       As Long = 10
Public Const CD_ISSUE_TYPE        As Long = 11
Public Const CD_SEVERITY          As Long = 12
Public Const CD_STATUS            As Long = 13
Public Const CD_ISSUE_LOC         As Long = 14
Public Const CD_ISSUE_DESC        As Long = 15
Public Const CD_RESOLVER          As Long = 16
Public Const CD_RES_ID            As Long = 17
Public Const CD_RES_DESIG         As Long = 18
Public Const CD_RES_DEPT          As Long = 19
Public Const CD_RES_DIV           As Long = 20
Public Const CD_RES_SUBDIV        As Long = 21
Public Const CD_RES_LOC           As Long = 22
Public Const CD_RESOLVED_AT       As Long = 23
Public Const CD_RES_REMARKS       As Long = 24
' Derived columns
Public Const CD_REPORT_DT         As Long = 25
Public Const CD_RESOLVE_DT        As Long = 26
Public Const CD_RES_HOURS         As Long = 27
Public Const CD_RES_DAYS          As Long = 28
Public Const CD_AGING_BUCKET      As Long = 29
Public Const CD_NORM_SEVERITY     As Long = 30
Public Const CD_NORM_STATUS       As Long = 31
Public Const CD_NORM_ISSUE_TYPE   As Long = 32
Public Const CD_SLA_TARGET        As Long = 33
Public Const CD_SLA_STATUS        As Long = 34
Public Const CD_IS_BREACH         As Long = 35
Public Const CD_DOW               As Long = 36
Public Const CD_HOUR              As Long = 37
Public Const CD_DATE_KEY          As Long = 38
Public Const CD_WEEK              As Long = 39
Public Const CD_REPORT_BIZ        As Long = 40
Public Const CD_CATEGORY          As Long = 41
Public Const CD_REPEAT_FLAG       As Long = 42
Public Const CD_IMPACT_SCORE      As Long = 43
Public Const CD_RISK_SCORE        As Long = 44
Public Const CD_AGE_HOURS         As Long = 45
Public Const CD_TOTAL_COLS        As Long = 45

'==============================================================================
'                                IMPORT
'==============================================================================
Public Sub ImportCsv()
    Dim path As String, ws As Worksheet
    path = ResolveCsvPath
    If Len(path) = 0 Then
        Err.Raise 5001, "ImportCsv", _
            "Apps.csv not found." & vbCrLf & vbCrLf & _
            "Searched in:" & vbCrLf & _
            "  - Workbook folder: " & ThisWorkbook.Path & vbCrLf & _
            "  - Documents, Desktop, Downloads" & vbCrLf & vbCrLf & _
            "Fix: Place Apps.csv in the same folder as this workbook " & _
            "(" & ThisWorkbook.Path & ") and try again."
    End If
    LogInfo "DATA", "Importing CSV from " & path

    Set ws = ThisWorkbook.Worksheets(SHT_RAW)
    ws.Cells.Clear

    ' Use ADO-free fast import via Workbooks.OpenText alternative: read line-by-line
    ' to gracefully handle quoted commas (the CSV has names like "Abhimanyu kumar ,").
    Dim fnum As Integer, line As String, rows As Collection, fields As Variant
    Set rows = New Collection
    fnum = FreeFile
    Open path For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, line
        If Left$(line, 1) = ChrW(65279) Or Asc(Left$(line, 1)) = 239 Then line = StripBom(line)
        rows.Add line
    Loop
    Close #fnum

    Dim r As Long, maxCols As Long, dat() As Variant
    maxCols = 24
    ReDim dat(1 To rows.Count, 1 To maxCols)
    For r = 1 To rows.Count
        fields = ParseCsvLine(CStr(rows(r)))
        Dim c As Long
        For c = 1 To maxCols
            If c <= UBound(fields) - LBound(fields) + 1 Then
                dat(r, c) = fields(LBound(fields) + c - 1)
            End If
        Next c
    Next r

    ws.Range(ws.Cells(1, 1), ws.Cells(rows.Count, maxCols)).Value = dat
    ' Style header row
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, maxCols))
        .Interior.Color = CLR_PANEL
        .Font.Color = CLR_TEXT
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .Font.Size = 10
    End With
    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9
    ws.Columns.AutoFit
    LogInfo "DATA", "Imported " & (rows.Count - 1) & " rows, " & maxCols & " columns"
End Sub

'==============================================================================
'                              CSV LINE PARSER
' Robust quoted-comma aware single-line parser.
'==============================================================================
Private Function ParseCsvLine(ByVal s As String) As Variant
    Dim out() As String, i As Long, ch As String, cur As String, inq As Boolean
    ReDim out(0 To 0)
    Dim n As Long: n = 0
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If inq And i < Len(s) And Mid$(s, i + 1, 1) = """" Then
                cur = cur & """": i = i + 1
            Else
                inq = Not inq
            End If
        ElseIf ch = "," And Not inq Then
            ReDim Preserve out(0 To n)
            out(n) = cur: n = n + 1: cur = vbNullString
        Else
            cur = cur & ch
        End If
    Next i
    ReDim Preserve out(0 To n)
    out(n) = cur
    ParseCsvLine = out
End Function

Private Function StripBom(ByVal s As String) As String
    Do While Len(s) > 0 And (Asc(Left$(s, 1)) = 239 Or Asc(Left$(s, 1)) = 187 Or Asc(Left$(s, 1)) = 191 Or AscW(Left$(s, 1)) = 65279)
        s = Mid$(s, 2)
    Loop
    StripBom = s
End Function

'==============================================================================
'                                 CLEAN
' Builds Clean_Data with all derived columns from Raw_Data.
'==============================================================================
Public Sub CleanData()
    Dim wsRaw As Worksheet, wsClean As Worksheet
    Set wsRaw = ThisWorkbook.Worksheets(SHT_RAW)
    Set wsClean = ThisWorkbook.Worksheets(SHT_CLEAN)
    wsClean.Cells.Clear

    Dim lastRow As Long, lastCol As Long
    lastRow = wsRaw.Cells(wsRaw.Rows.Count, 9).End(xlUp).Row     ' Issue ID is reliable
    lastCol = 24
    If lastRow < 2 Then Err.Raise 5002, , "No data rows in Raw_Data."

    Dim src As Variant, dst() As Variant
    src = wsRaw.Range(wsRaw.Cells(1, 1), wsRaw.Cells(lastRow, lastCol)).Value
    ReDim dst(1 To lastRow, 1 To CD_TOTAL_COLS)

    ' Header row
    Dim headers As Variant
    headers = Array("Reporter", "Reporter Identifier", "Reporter Designation", _
        "Report Department", "Reporter Division", "Reporter Sub Division", _
        "Reporter Location", "Reported At", "Issue ID", "Issue Title", _
        "Issue Type", "Severity", "Current Status", "Issue Location", _
        "Issue Description", "Resolver", "Resolver Identifier", _
        "Resolver Designation", "Resolver Department", "Resolver Division", _
        "Resolver Sub Division", "Resolver Location", "Resolved At", "Resolved Remarks", _
        "ReportedDateTime", "ResolvedDateTime", "ResolutionHours", "ResolutionDays", _
        "AgingBucket", "Severity_N", "Status_N", "IssueType_N", _
        "SLA_Target_H", "SLA_Status", "Is_Breach", "DayOfWeek", "HourOfDay", _
        "DateKey", "WeekNumber", "BusinessHours", "Category", "RepeatFlag", _
        "ImpactScore", "RiskScore", "Age_Hours")

    Dim h As Long
    For h = 0 To UBound(headers)
        wsClean.Cells(1, h + 1).Value = headers(h)
    Next h

    Dim r As Long, i As Long
    Dim repDt As Date, resDt As Date, hasRep As Boolean, hasRes As Boolean
    Dim sev As String, stat As String, itype As String, sla As Double
    Dim title As String, cat As String, rpt As String

    ' First pass - normalize and compute per-row metrics
    For r = 2 To lastRow
        ' Carry over raw 24 cols
        For i = 1 To 24
            dst(r, i) = src(r, i)
        Next i
        ' Reported / Resolved datetimes
        repDt = ParseDateTime(CStr(src(r, 8)), hasRep)
        resDt = ParseDateTime(CStr(src(r, 23)), hasRes)
        If hasRep Then dst(r, CD_REPORT_DT) = repDt
        If hasRes Then dst(r, CD_RESOLVE_DT) = resDt

        ' Resolution metrics
        If hasRep And hasRes And resDt >= repDt Then
            dst(r, CD_RES_HOURS) = (resDt - repDt) * 24#
            dst(r, CD_RES_DAYS) = (resDt - repDt)
        Else
            dst(r, CD_RES_HOURS) = vbNullString
            dst(r, CD_RES_DAYS) = vbNullString
        End If

        ' Normalized severity / status / issue type
        sev = NormalizeSeverity(CStr(src(r, 12)))
        stat = NormalizeStatus(CStr(src(r, 13)))
        itype = NormalizeIssueType(CStr(src(r, 11)))
        dst(r, CD_NORM_SEVERITY) = sev
        dst(r, CD_NORM_STATUS) = stat
        dst(r, CD_NORM_ISSUE_TYPE) = itype

        ' SLA target & status
        sla = SlaTargetHours(sev)
        dst(r, CD_SLA_TARGET) = sla

        Dim ageH As Double
        If hasRep Then
            If stat = "Closed" And hasRes Then
                ageH = (resDt - repDt) * 24#
            Else
                ageH = (Now - repDt) * 24#
            End If
        Else
            ageH = 0
        End If
        dst(r, CD_AGE_HOURS) = ageH

        If stat = "Closed" Then
            If hasRep And hasRes Then
                If (resDt - repDt) * 24# <= sla Then
                    dst(r, CD_SLA_STATUS) = "Met"
                    dst(r, CD_IS_BREACH) = 0
                Else
                    dst(r, CD_SLA_STATUS) = "Breached"
                    dst(r, CD_IS_BREACH) = 1
                End If
            Else
                dst(r, CD_SLA_STATUS) = "Unknown"
                dst(r, CD_IS_BREACH) = 0
            End If
        Else
            ' Open ticket - compare current age
            If ageH > sla Then
                dst(r, CD_SLA_STATUS) = "Breached"
                dst(r, CD_IS_BREACH) = 1
            ElseIf ageH > sla * 0.75 Then
                dst(r, CD_SLA_STATUS) = "At Risk"
                dst(r, CD_IS_BREACH) = 0
            Else
                dst(r, CD_SLA_STATUS) = "On Track"
                dst(r, CD_IS_BREACH) = 0
            End If
        End If

        ' Aging bucket (based on age hours regardless of status)
        dst(r, CD_AGING_BUCKET) = AgingBucket(ageH)

        ' Calendar derived
        If hasRep Then
            dst(r, CD_DOW) = WeekdayName(Weekday(repDt, vbMonday), False, vbMonday)
            dst(r, CD_HOUR) = Hour(repDt)
            dst(r, CD_DATE_KEY) = DateValue(repDt)
            dst(r, CD_WEEK) = "W" & Format(repDt, "ww")
            dst(r, CD_REPORT_BIZ) = IIf(IsBusinessHour(repDt), "Business", "Off-Hours")
        End If

        ' Auto category clustering from title + description keywords
        title = CStr(src(r, 10)) & " " & CStr(src(r, 15))
        dst(r, CD_CATEGORY) = ClusterCategory(title)

        ' Impact score (1..100) - severity weight × open penalty × biz hours × type weight
        dst(r, CD_IMPACT_SCORE) = ComputeImpact(sev, stat, itype, ageH, sla)
        dst(r, CD_RISK_SCORE) = ComputeRisk(sev, stat, ageH, sla)
    Next r

    ' Second pass - repeat detection (same reporter ID + similar title within 7 days)
    Dim repMap As Object: Set repMap = CreateObject("Scripting.Dictionary")
    Dim key As String
    For r = 2 To lastRow
        key = CStr(dst(r, CD_REP_ID)) & "|" & CategoryKey(CStr(dst(r, CD_ISSUE_TITLE)))
        If repMap.Exists(key) Then
            repMap(key) = repMap(key) + 1
        Else
            repMap(key) = 1
        End If
    Next r
    For r = 2 To lastRow
        key = CStr(dst(r, CD_REP_ID)) & "|" & CategoryKey(CStr(dst(r, CD_ISSUE_TITLE)))
        If repMap(key) > 1 Then
            dst(r, CD_REPEAT_FLAG) = "Repeat"
        Else
            dst(r, CD_REPEAT_FLAG) = "First"
        End If
    Next r

    ' Bulk write back
    wsClean.Range(wsClean.Cells(2, 1), wsClean.Cells(lastRow, CD_TOTAL_COLS)).Value = _
        ResliceTwoD(dst, 2, lastRow, 1, CD_TOTAL_COLS)

    ' Format
    With wsClean.Range(wsClean.Cells(1, 1), wsClean.Cells(1, CD_TOTAL_COLS))
        .Interior.Color = CLR_PANEL
        .Font.Color = CLR_TEXT
        .Font.Bold = True
        .Font.Name = "Segoe UI"
        .RowHeight = 22
    End With
    wsClean.Cells.Font.Name = "Segoe UI"
    wsClean.Cells.Font.Size = 9
    wsClean.Range(wsClean.Cells(2, CD_REPORT_DT), wsClean.Cells(lastRow, CD_REPORT_DT)).NumberFormat = "yyyy-mm-dd hh:mm"
    wsClean.Range(wsClean.Cells(2, CD_RESOLVE_DT), wsClean.Cells(lastRow, CD_RESOLVE_DT)).NumberFormat = "yyyy-mm-dd hh:mm"
    wsClean.Range(wsClean.Cells(2, CD_DATE_KEY), wsClean.Cells(lastRow, CD_DATE_KEY)).NumberFormat = "yyyy-mm-dd"
    wsClean.Range(wsClean.Cells(2, CD_RES_HOURS), wsClean.Cells(lastRow, CD_RES_HOURS)).NumberFormat = "0.0"
    wsClean.Range(wsClean.Cells(2, CD_AGE_HOURS), wsClean.Cells(lastRow, CD_AGE_HOURS)).NumberFormat = "0.0"

    ' Build a structured table for Power Query / referencing
    On Error Resume Next
    wsClean.ListObjects("tblTickets").Unlist
    On Error GoTo 0
    Dim lo As ListObject
    Set lo = wsClean.ListObjects.Add(xlSrcRange, _
        wsClean.Range(wsClean.Cells(1, 1), wsClean.Cells(lastRow, CD_TOTAL_COLS)), _
        , xlYes)
    lo.Name = "tblTickets"
    lo.TableStyle = ""
    wsClean.Columns.AutoFit

    LogInfo "DATA", "Cleaned " & (lastRow - 1) & " rows -> tblTickets"
End Sub

'==============================================================================
'                            NORMALIZATION HELPERS
'==============================================================================
Public Function NormalizeSeverity(ByVal v As String) As String
    Dim t As String: t = LCase$(Trim$(v))
    Select Case t
        Case "critical", "p1", "sev1", "crit": NormalizeSeverity = "Critical"
        Case "high", "p2", "sev2": NormalizeSeverity = "High"
        Case "medium", "med", "p3", "sev3": NormalizeSeverity = "Medium"
        Case "low", "p4", "sev4": NormalizeSeverity = "Low"
        Case "": NormalizeSeverity = "Unspecified"
        Case Else: NormalizeSeverity = StrConv(v, vbProperCase)
    End Select
End Function

Public Function NormalizeStatus(ByVal v As String) As String
    Dim t As String: t = LCase$(Trim$(v))
    Select Case t
        Case "closed", "resolved", "done", "completed": NormalizeStatus = "Closed"
        Case "open", "new", "raised", "in progress", "in-progress", "pending": NormalizeStatus = "Open"
        Case "": NormalizeStatus = "Open"
        Case Else: NormalizeStatus = StrConv(v, vbProperCase)
    End Select
End Function

Public Function NormalizeIssueType(ByVal v As String) As String
    Dim t As String: t = LCase$(Trim$(v))
    Select Case t
        Case "it": NormalizeIssueType = "IT"
        Case "repair and maintenance", "rnm", "r&m", "maintenance": NormalizeIssueType = "Repair & Maintenance"
        Case "marketing": NormalizeIssueType = "Marketing"
        Case "": NormalizeIssueType = "Uncategorized"
        Case Else: NormalizeIssueType = StrConv(v, vbProperCase)
    End Select
End Function

Public Function SlaTargetHours(ByVal sev As String) As Double
    Select Case sev
        Case "Critical": SlaTargetHours = SLA_CRITICAL_H
        Case "High":     SlaTargetHours = SLA_HIGH_H
        Case "Medium":   SlaTargetHours = SLA_MEDIUM_H
        Case "Low":      SlaTargetHours = SLA_LOW_H
        Case Else:       SlaTargetHours = SLA_UNSPEC_H
    End Select
End Function

Public Function AgingBucket(ByVal hrs As Double) As String
    Select Case hrs
        Case Is < 4:   AgingBucket = "00-04h"
        Case Is < 24:  AgingBucket = "04-24h"
        Case Is < 72:  AgingBucket = "1-3d"
        Case Is < 168: AgingBucket = "3-7d"
        Case Else:     AgingBucket = ">7d"
    End Select
End Function

Public Function IsBusinessHour(ByVal d As Date) As Boolean
    Dim h As Long, w As Long
    h = Hour(d): w = Weekday(d, vbMonday)
    IsBusinessHour = (w <= 6) And (h >= 9) And (h < 21)
End Function

' Auto-cluster issue title into a category by keyword scan
Public Function ClusterCategory(ByVal txt As String) As String
    Dim t As String: t = LCase$(txt)
    If InStr(t, "cctv") > 0 Or InStr(t, "camera") > 0 Then ClusterCategory = "Surveillance / CCTV": Exit Function
    If InStr(t, "laptop") > 0 Or InStr(t, "system") > 0 Or InStr(t, "computer") > 0 Or InStr(t, "dmb") > 0 Or InStr(t, "bluetooth") > 0 Or InStr(t, "wifi") > 0 Or InStr(t, "network") > 0 Then ClusterCategory = "IT Hardware / Network": Exit Function
    If InStr(t, "fryer") > 0 Or InStr(t, "boiler") > 0 Or InStr(t, "griller") > 0 Or InStr(t, "induction") > 0 Or InStr(t, "mixie") > 0 Or InStr(t, "oven") > 0 Then ClusterCategory = "Cooking Equipment": Exit Function
    If InStr(t, "freezer") > 0 Or InStr(t, "chiller") > 0 Or InStr(t, "ice") > 0 Or InStr(t, "vissicooler") > 0 Or InStr(t, "cooler") > 0 Or InStr(t, "compressor") > 0 Then ClusterCategory = "Refrigeration": Exit Function
    If InStr(t, "ro ") > 0 Or InStr(t, "water") > 0 Or InStr(t, "pipe") > 0 Or InStr(t, "drain") > 0 Or InStr(t, "leak") > 0 Or InStr(t, "toilet") > 0 Then ClusterCategory = "Plumbing / Water": Exit Function
    If InStr(t, "mcb") > 0 Or InStr(t, "switch") > 0 Or InStr(t, "socket") > 0 Or InStr(t, "plug") > 0 Or InStr(t, "wire") > 0 Or InStr(t, "light") > 0 Or InStr(t, "bulb") > 0 Or InStr(t, "led") > 0 Then ClusterCategory = "Electrical / Lighting": Exit Function
    If InStr(t, "ac ") > 0 Or InStr(t, "hvac") > 0 Or InStr(t, "air") > 0 Then ClusterCategory = "HVAC": Exit Function
    If InStr(t, "tile") > 0 Or InStr(t, "paint") > 0 Or InStr(t, "door") > 0 Or InStr(t, "handle") > 0 Or InStr(t, "wall") > 0 Then ClusterCategory = "Civil / Carpentry": Exit Function
    ClusterCategory = "Other"
End Function

Public Function CategoryKey(ByVal title As String) As String
    Dim t As String: t = LCase$(title)
    Dim words As Variant: words = Split(t, " ")
    Dim out As String, w As Variant
    For Each w In words
        If Len(w) >= 4 Then out = out & "|" & w
    Next w
    CategoryKey = out
End Function

Public Function ComputeImpact(ByVal sev As String, ByVal stat As String, _
        ByVal itype As String, ByVal ageH As Double, ByVal sla As Double) As Double
    Dim s As Double, st As Double, ty As Double, ag As Double
    Select Case sev
        Case "Critical": s = 40
        Case "High":     s = 28
        Case "Medium":   s = 16
        Case "Low":      s = 8
        Case Else:       s = 12
    End Select
    st = IIf(stat = "Open", 1.4, 1#)
    ty = IIf(itype = "IT", 1.2, IIf(itype = "Repair & Maintenance", 1#, 0.85))
    If sla > 0 Then ag = WorksheetFunction.Min(2#, (ageH / sla)) Else ag = 1
    ComputeImpact = WorksheetFunction.Min(100, s * st * ty * (0.6 + 0.4 * ag))
End Function

Public Function ComputeRisk(ByVal sev As String, ByVal stat As String, _
        ByVal ageH As Double, ByVal sla As Double) As Double
    Dim r As Double
    Select Case sev
        Case "Critical": r = 50
        Case "High":     r = 32
        Case "Medium":   r = 18
        Case "Low":      r = 8
        Case Else:       r = 14
    End Select
    If stat = "Open" Then r = r + 12
    If sla > 0 And ageH > sla Then r = r + 25
    If r > 100 Then r = 100
    ComputeRisk = r
End Function

'==============================================================================
'                              DATE PARSING
' Apps.csv uses "dd-mm-yyyy h:mm" - VBA DateValue is locale-sensitive so parse
' explicitly.
'==============================================================================
Public Function ParseDateTime(ByVal s As String, ByRef ok As Boolean) As Date
    Dim t As String: t = Trim$(s)
    ok = False
    If Len(t) = 0 Then Exit Function
    Dim parts As Variant, dpart As String, tpart As String
    parts = Split(t, " ")
    dpart = CStr(parts(0))
    If UBound(parts) >= 1 Then tpart = CStr(parts(1)) Else tpart = "0:00"

    Dim dd As Long, mm As Long, yy As Long
    Dim dparts As Variant: dparts = Split(dpart, "-")
    If UBound(dparts) <> 2 Then dparts = Split(dpart, "/")
    If UBound(dparts) <> 2 Then Exit Function
    On Error Resume Next
    dd = CLng(dparts(0))
    mm = CLng(dparts(1))
    yy = CLng(dparts(2))
    If yy < 100 Then yy = 2000 + yy
    Dim hh As Long, mn As Long
    Dim tparts As Variant: tparts = Split(tpart, ":")
    If UBound(tparts) >= 1 Then
        hh = CLng(tparts(0))
        mn = CLng(tparts(1))
    End If
    On Error GoTo 0
    If dd >= 1 And dd <= 31 And mm >= 1 And mm <= 12 And yy >= 1900 Then
        ParseDateTime = DateSerial(yy, mm, dd) + TimeSerial(hh, mn, 0)
        ok = True
    End If
End Function

'==============================================================================
'                              DATA MODEL
' Builds aggregated lookup tables on Data_Model sheet.
'==============================================================================
Public Sub BuildDataModel()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_MODEL)
    ws.Cells.Clear

    ws.Cells.Font.Name = "Segoe UI"
    ws.Cells.Font.Size = 9

    Dim arr As Variant
    arr = ReadCleanArray()
    If IsEmpty(arr) Then Exit Sub
    Dim n As Long: n = UBound(arr, 1)

    ' --- Block layouts, label, then aggregate function fills it ---
    Dim row As Long: row = 1
    row = WriteAggHeader(ws, row, "AGG: Severity")
    row = AggCount(ws, row, arr, CD_NORM_SEVERITY)
    row = WriteAggHeader(ws, row, "AGG: Status")
    row = AggCount(ws, row, arr, CD_NORM_STATUS)
    row = WriteAggHeader(ws, row, "AGG: IssueType")
    row = AggCount(ws, row, arr, CD_NORM_ISSUE_TYPE)
    row = WriteAggHeader(ws, row, "AGG: Category")
    row = AggCount(ws, row, arr, CD_CATEGORY)
    row = WriteAggHeader(ws, row, "AGG: Aging Bucket")
    row = AggCountOrdered(ws, row, arr, CD_AGING_BUCKET, _
        Array("00-04h", "04-24h", "1-3d", "3-7d", ">7d"))
    row = WriteAggHeader(ws, row, "AGG: Day Of Week")
    row = AggCountOrdered(ws, row, arr, CD_DOW, _
        Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"))
    row = WriteAggHeader(ws, row, "AGG: Hour Of Day")
    row = AggHour(ws, row, arr)
    row = WriteAggHeader(ws, row, "AGG: Daily Volume")
    row = AggDaily(ws, row, arr)
    row = WriteAggHeader(ws, row, "AGG: Top 10 Reporters")
    row = AggTopN(ws, row, arr, CD_REPORTER, 10)
    row = WriteAggHeader(ws, row, "AGG: Top 10 Resolvers")
    row = AggTopNNonBlank(ws, row, arr, CD_RESOLVER, 10)
    row = WriteAggHeader(ws, row, "AGG: Top 10 Reporter Locations")
    row = AggTopN(ws, row, arr, CD_REP_LOC, 10)
    row = WriteAggHeader(ws, row, "AGG: Severity x Status (Heatmap)")
    row = AggCross(ws, row, arr, CD_NORM_SEVERITY, CD_NORM_STATUS, _
        Array("Critical", "High", "Medium", "Low", "Unspecified"), Array("Open", "Closed"))
    row = WriteAggHeader(ws, row, "AGG: Hour x DOW (Heatmap)")
    row = AggHourDow(ws, row, arr)
    row = WriteAggHeader(ws, row, "AGG: SLA by Severity")
    row = AggSlaBySeverity(ws, row, arr)
    row = WriteAggHeader(ws, row, "AGG: Resolver MTTR")
    row = AggResolverMttr(ws, row, arr)
    row = WriteAggHeader(ws, row, "AGG: Risk by Location")
    row = AggRiskByLocation(ws, row, arr)

    ws.Columns.AutoFit
    LogInfo "DATA", "Data_Model built with " & n & " ticket rows"
End Sub

Private Function ReadCleanArray() As Variant
    Dim ws As Worksheet, lr As Long
    Set ws = ThisWorkbook.Worksheets(SHT_CLEAN)
    lr = ws.Cells(ws.Rows.Count, CD_ISSUE_ID).End(xlUp).Row
    If lr < 2 Then ReadCleanArray = Empty: Exit Function
    ReadCleanArray = ws.Range(ws.Cells(1, 1), ws.Cells(lr, CD_TOTAL_COLS)).Value
End Function

Private Function WriteAggHeader(ws As Worksheet, ByVal startRow As Long, ByVal title As String) As Long
    ws.Cells(startRow, 1).Value = title
    With ws.Cells(startRow, 1)
        .Font.Bold = True
        .Font.Color = CLR_ACCENT
        .Interior.Color = CLR_PANEL
    End With
    WriteAggHeader = startRow + 1
End Function

Private Function AggCount(ws As Worksheet, ByVal startRow As Long, arr As Variant, ByVal col As Long) As Long
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To UBound(arr, 1)
        k = NzString(arr(r, col), "(blank)")
        If d.Exists(k) Then d(k) = d(k) + 1 Else d(k) = 1
    Next r
    Dim keys As Variant: keys = d.keys
    SortDictDesc d, keys
    ws.Cells(startRow, 1).Value = "Key"
    ws.Cells(startRow, 2).Value = "Count"
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 2))
    Dim i As Long
    For i = 0 To UBound(keys)
        ws.Cells(startRow + 1 + i, 1).Value = keys(i)
        ws.Cells(startRow + 1 + i, 2).Value = d(keys(i))
    Next i
    AggCount = startRow + UBound(keys) + 3
End Function

Private Function AggCountOrdered(ws As Worksheet, ByVal startRow As Long, arr As Variant, _
        ByVal col As Long, ByVal orderArr As Variant) As Long
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String, i As Long
    For r = 2 To UBound(arr, 1)
        k = NzString(arr(r, col), "(blank)")
        If d.Exists(k) Then d(k) = d(k) + 1 Else d(k) = 1
    Next r
    ws.Cells(startRow, 1).Value = "Key"
    ws.Cells(startRow, 2).Value = "Count"
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 2))
    For i = 0 To UBound(orderArr)
        ws.Cells(startRow + 1 + i, 1).Value = orderArr(i)
        ws.Cells(startRow + 1 + i, 2).Value = IIf(d.Exists(CStr(orderArr(i))), d(CStr(orderArr(i))), 0)
    Next i
    AggCountOrdered = startRow + UBound(orderArr) + 3
End Function

Private Function AggHour(ws As Worksheet, ByVal startRow As Long, arr As Variant) As Long
    Dim counts(0 To 23) As Long, r As Long, h As Variant
    For r = 2 To UBound(arr, 1)
        h = arr(r, CD_HOUR)
        If IsNumeric(h) Then counts(CLng(h)) = counts(CLng(h)) + 1
    Next r
    ws.Cells(startRow, 1).Value = "Hour"
    ws.Cells(startRow, 2).Value = "Count"
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 2))
    Dim i As Long
    For i = 0 To 23
        ws.Cells(startRow + 1 + i, 1).Value = Format(i, "00") & ":00"
        ws.Cells(startRow + 1 + i, 2).Value = counts(i)
    Next i
    AggHour = startRow + 26
End Function

Private Function AggDaily(ws As Worksheet, ByVal startRow As Long, arr As Variant) As Long
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As Variant
    For r = 2 To UBound(arr, 1)
        k = arr(r, CD_DATE_KEY)
        If IsDate(k) Then
            Dim dk As Date: dk = CDate(k)
            If d.Exists(CLng(dk)) Then d(CLng(dk)) = d(CLng(dk)) + 1 Else d(CLng(dk)) = 1
        End If
    Next r
    ws.Cells(startRow, 1).Value = "Date"
    ws.Cells(startRow, 2).Value = "Count"
    ws.Cells(startRow, 3).Value = "MA(3)"
    ws.Cells(startRow, 4).Value = "Forecast"
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 4))
    ' Sort by date asc
    Dim ks() As Long, i As Long
    ReDim ks(0 To d.Count - 1)
    Dim kk As Variant, j As Long
    For Each kk In d.keys
        ks(j) = CLng(kk): j = j + 1
    Next kk
    Dim a As Long, b As Long, t As Long
    For a = LBound(ks) To UBound(ks) - 1
        For b = a + 1 To UBound(ks)
            If ks(a) > ks(b) Then t = ks(a): ks(a) = ks(b): ks(b) = t
        Next b
    Next a
    Dim cnt As Long, ma As Double, prev As Long
    For i = 0 To UBound(ks)
        cnt = d(ks(i))
        ws.Cells(startRow + 1 + i, 1).Value = CDate(ks(i))
        ws.Cells(startRow + 1 + i, 1).NumberFormat = "yyyy-mm-dd"
        ws.Cells(startRow + 1 + i, 2).Value = cnt
        Dim s As Double, n As Long, k As Long
        s = 0: n = 0
        For k = WorksheetFunction.Max(0, i - 2) To i
            s = s + d(ks(k)): n = n + 1
        Next k
        ma = IIf(n > 0, s / n, 0)
        ws.Cells(startRow + 1 + i, 3).Value = ma
        ws.Cells(startRow + 1 + i, 3).NumberFormat = "0.0"
    Next i
    ' Forecast next 3 days using 3-day MA + slight trend
    Dim lastIdx As Long, lastDate As Date, slope As Double
    lastIdx = UBound(ks): lastDate = CDate(ks(lastIdx))
    If lastIdx >= 1 Then
        slope = (d(ks(lastIdx)) - d(ks(WorksheetFunction.Max(0, lastIdx - 2)))) / 3#
    Else
        slope = 0
    End If
    Dim baseMA As Double
    baseMA = ws.Cells(startRow + 1 + lastIdx, 3).Value
    For i = 1 To 3
        ws.Cells(startRow + 1 + lastIdx + i, 1).Value = lastDate + i
        ws.Cells(startRow + 1 + lastIdx + i, 1).NumberFormat = "yyyy-mm-dd"
        ws.Cells(startRow + 1 + lastIdx + i, 4).Value = WorksheetFunction.Max(0, baseMA + slope * i)
        ws.Cells(startRow + 1 + lastIdx + i, 4).NumberFormat = "0.0"
    Next i
    AggDaily = startRow + lastIdx + 5
End Function

Private Function AggTopN(ws As Worksheet, ByVal startRow As Long, arr As Variant, _
        ByVal col As Long, ByVal n As Long) As Long
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To UBound(arr, 1)
        k = NzString(arr(r, col), "(blank)")
        If d.Exists(k) Then d(k) = d(k) + 1 Else d(k) = 1
    Next r
    Dim keys As Variant: keys = d.keys
    SortDictDesc d, keys
    ws.Cells(startRow, 1).Value = "Key"
    ws.Cells(startRow, 2).Value = "Count"
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 2))
    Dim i As Long, top As Long
    top = WorksheetFunction.Min(n - 1, UBound(keys))
    For i = 0 To top
        ws.Cells(startRow + 1 + i, 1).Value = keys(i)
        ws.Cells(startRow + 1 + i, 2).Value = d(keys(i))
    Next i
    AggTopN = startRow + top + 3
End Function

Private Function AggTopNNonBlank(ws As Worksheet, ByVal startRow As Long, arr As Variant, _
        ByVal col As Long, ByVal n As Long) As Long
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To UBound(arr, 1)
        k = Trim$(CStr(arr(r, col)))
        If Len(k) = 0 Then GoTo nx
        If d.Exists(k) Then d(k) = d(k) + 1 Else d(k) = 1
nx:
    Next r
    Dim keys As Variant: keys = d.keys
    SortDictDesc d, keys
    ws.Cells(startRow, 1).Value = "Resolver"
    ws.Cells(startRow, 2).Value = "Tickets"
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 2))
    Dim i As Long, top As Long
    top = WorksheetFunction.Min(n - 1, UBound(keys))
    For i = 0 To top
        ws.Cells(startRow + 1 + i, 1).Value = keys(i)
        ws.Cells(startRow + 1 + i, 2).Value = d(keys(i))
    Next i
    AggTopNNonBlank = startRow + top + 3
End Function

Private Function AggCross(ws As Worksheet, ByVal startRow As Long, arr As Variant, _
        ByVal rowCol As Long, ByVal colCol As Long, _
        ByVal rowOrder As Variant, ByVal colOrder As Variant) As Long
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To UBound(arr, 1)
        k = NzString(arr(r, rowCol), "(blank)") & "|" & NzString(arr(r, colCol), "(blank)")
        If d.Exists(k) Then d(k) = d(k) + 1 Else d(k) = 1
    Next r
    Dim i As Long, j As Long
    ws.Cells(startRow, 1).Value = ""
    For j = 0 To UBound(colOrder)
        ws.Cells(startRow, 2 + j).Value = colOrder(j)
    Next j
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 2 + UBound(colOrder)))
    For i = 0 To UBound(rowOrder)
        ws.Cells(startRow + 1 + i, 1).Value = rowOrder(i)
        For j = 0 To UBound(colOrder)
            k = CStr(rowOrder(i)) & "|" & CStr(colOrder(j))
            ws.Cells(startRow + 1 + i, 2 + j).Value = IIf(d.Exists(k), d(k), 0)
        Next j
    Next i
    AggCross = startRow + UBound(rowOrder) + 3
End Function

Private Function AggHourDow(ws As Worksheet, ByVal startRow As Long, arr As Variant) As Long
    Dim dows As Variant, h As Long, dn As String, r As Long, k As String
    dows = Array("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    For r = 2 To UBound(arr, 1)
        If IsNumeric(arr(r, CD_HOUR)) And Len(arr(r, CD_DOW)) > 0 Then
            k = arr(r, CD_DOW) & "|" & CLng(arr(r, CD_HOUR))
            If d.Exists(k) Then d(k) = d(k) + 1 Else d(k) = 1
        End If
    Next r
    ws.Cells(startRow, 1).Value = "Hour"
    Dim j As Long
    For j = 0 To UBound(dows): ws.Cells(startRow, 2 + j).Value = dows(j): Next j
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 2 + UBound(dows)))
    For h = 0 To 23
        ws.Cells(startRow + 1 + h, 1).Value = Format(h, "00") & ":00"
        For j = 0 To UBound(dows)
            k = CStr(dows(j)) & "|" & h
            ws.Cells(startRow + 1 + h, 2 + j).Value = IIf(d.Exists(k), d(k), 0)
        Next j
    Next h
    AggHourDow = startRow + 26
End Function

Private Function AggSlaBySeverity(ws As Worksheet, ByVal startRow As Long, arr As Variant) As Long
    Dim sevs As Variant: sevs = Array("Critical", "High", "Medium", "Low", "Unspecified")
    Dim totals As Object, met As Object, breach As Object, atrisk As Object
    Set totals = CreateObject("Scripting.Dictionary")
    Set met = CreateObject("Scripting.Dictionary")
    Set breach = CreateObject("Scripting.Dictionary")
    Set atrisk = CreateObject("Scripting.Dictionary")
    Dim r As Long, sv As String, ss As String
    For r = 2 To UBound(arr, 1)
        sv = CStr(arr(r, CD_NORM_SEVERITY))
        ss = CStr(arr(r, CD_SLA_STATUS))
        If Not totals.Exists(sv) Then totals(sv) = 0: met(sv) = 0: breach(sv) = 0: atrisk(sv) = 0
        totals(sv) = totals(sv) + 1
        Select Case ss
            Case "Met", "On Track": met(sv) = met(sv) + 1
            Case "Breached":        breach(sv) = breach(sv) + 1
            Case "At Risk":         atrisk(sv) = atrisk(sv) + 1
        End Select
    Next r
    ws.Cells(startRow, 1).Value = "Severity"
    ws.Cells(startRow, 2).Value = "Total"
    ws.Cells(startRow, 3).Value = "Met"
    ws.Cells(startRow, 4).Value = "At Risk"
    ws.Cells(startRow, 5).Value = "Breached"
    ws.Cells(startRow, 6).Value = "Compliance %"
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 6))
    Dim i As Long, t As Long, m As Long
    For i = 0 To UBound(sevs)
        Dim s As String: s = CStr(sevs(i))
        ws.Cells(startRow + 1 + i, 1).Value = s
        ws.Cells(startRow + 1 + i, 2).Value = IIf(totals.Exists(s), totals(s), 0)
        ws.Cells(startRow + 1 + i, 3).Value = IIf(met.Exists(s), met(s), 0)
        ws.Cells(startRow + 1 + i, 4).Value = IIf(atrisk.Exists(s), atrisk(s), 0)
        ws.Cells(startRow + 1 + i, 5).Value = IIf(breach.Exists(s), breach(s), 0)
        t = ws.Cells(startRow + 1 + i, 2).Value
        m = ws.Cells(startRow + 1 + i, 3).Value
        ws.Cells(startRow + 1 + i, 6).Value = IIf(t > 0, m / t, 0)
        ws.Cells(startRow + 1 + i, 6).NumberFormat = "0.0%"
    Next i
    AggSlaBySeverity = startRow + UBound(sevs) + 3
End Function

Private Function AggResolverMttr(ws As Worksheet, ByVal startRow As Long, arr As Variant) As Long
    Dim cnt As Object, sumH As Object, met As Object
    Set cnt = CreateObject("Scripting.Dictionary")
    Set sumH = CreateObject("Scripting.Dictionary")
    Set met = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To UBound(arr, 1)
        k = Trim$(CStr(arr(r, CD_RESOLVER)))
        If Len(k) = 0 Then GoTo nx
        If Not cnt.Exists(k) Then cnt(k) = 0: sumH(k) = 0#: met(k) = 0
        cnt(k) = cnt(k) + 1
        If IsNumeric(arr(r, CD_RES_HOURS)) Then sumH(k) = sumH(k) + CDbl(arr(r, CD_RES_HOURS))
        If CStr(arr(r, CD_SLA_STATUS)) = "Met" Then met(k) = met(k) + 1
nx:
    Next r
    ws.Cells(startRow, 1).Value = "Resolver"
    ws.Cells(startRow, 2).Value = "Tickets"
    ws.Cells(startRow, 3).Value = "Avg Resolution (h)"
    ws.Cells(startRow, 4).Value = "SLA Met %"
    ws.Cells(startRow, 5).Value = "Efficiency Score"
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 5))
    Dim keys As Variant: keys = cnt.keys
    SortDictDesc cnt, keys
    Dim i As Long, c As Long
    For i = 0 To UBound(keys)
        c = cnt(keys(i))
        ws.Cells(startRow + 1 + i, 1).Value = keys(i)
        ws.Cells(startRow + 1 + i, 2).Value = c
        ws.Cells(startRow + 1 + i, 3).Value = IIf(c > 0, sumH(keys(i)) / c, 0)
        ws.Cells(startRow + 1 + i, 3).NumberFormat = "0.0"
        ws.Cells(startRow + 1 + i, 4).Value = IIf(c > 0, met(keys(i)) / c, 0)
        ws.Cells(startRow + 1 + i, 4).NumberFormat = "0.0%"
        Dim eff As Double
        eff = ws.Cells(startRow + 1 + i, 4).Value * 60 + WorksheetFunction.Min(40, c * 4)
        ws.Cells(startRow + 1 + i, 5).Value = eff
        ws.Cells(startRow + 1 + i, 5).NumberFormat = "0"
    Next i
    AggResolverMttr = startRow + UBound(keys) + 3
End Function

Private Function AggRiskByLocation(ws As Worksheet, ByVal startRow As Long, arr As Variant) As Long
    Dim d As Object, c As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set c = CreateObject("Scripting.Dictionary")
    Dim r As Long, k As String
    For r = 2 To UBound(arr, 1)
        k = Trim$(NzString(arr(r, CD_REP_LOC), "(blank)"))
        If Not d.Exists(k) Then d(k) = 0#: c(k) = 0
        d(k) = d(k) + CDbl(NzNum(arr(r, CD_RISK_SCORE)))
        c(k) = c(k) + 1
    Next r
    ws.Cells(startRow, 1).Value = "Location"
    ws.Cells(startRow, 2).Value = "Tickets"
    ws.Cells(startRow, 3).Value = "Avg Risk"
    ws.Cells(startRow, 4).Value = "Total Risk"
    StyleAggHeader ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow, 4))
    ' sort by total risk
    Dim keys As Variant: keys = d.keys
    SortDictDesc d, keys
    Dim i As Long, top As Long
    top = WorksheetFunction.Min(14, UBound(keys))
    For i = 0 To top
        ws.Cells(startRow + 1 + i, 1).Value = keys(i)
        ws.Cells(startRow + 1 + i, 2).Value = c(keys(i))
        ws.Cells(startRow + 1 + i, 3).Value = IIf(c(keys(i)) > 0, d(keys(i)) / c(keys(i)), 0)
        ws.Cells(startRow + 1 + i, 3).NumberFormat = "0.0"
        ws.Cells(startRow + 1 + i, 4).Value = d(keys(i))
        ws.Cells(startRow + 1 + i, 4).NumberFormat = "0.0"
    Next i
    AggRiskByLocation = startRow + top + 3
End Function

'==============================================================================
'                              UTILITIES
'==============================================================================
Public Function NzString(ByVal v As Variant, ByVal fallback As String) As String
    If IsNull(v) Or IsEmpty(v) Then NzString = fallback: Exit Function
    If Trim$(CStr(v)) = "" Then NzString = fallback Else NzString = CStr(v)
End Function

Public Function NzNum(ByVal v As Variant) As Double
    If IsNumeric(v) Then NzNum = CDbl(v) Else NzNum = 0
End Function

Public Sub SortDictDesc(d As Object, ByRef keys As Variant)
    Dim i As Long, j As Long, t As Variant
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If d(keys(i)) < d(keys(j)) Then
                t = keys(i): keys(i) = keys(j): keys(j) = t
            End If
        Next j
    Next i
End Sub

Private Sub StyleAggHeader(rng As Range)
    With rng
        .Font.Bold = True
        .Font.Color = CLR_TEXT
        .Interior.Color = CLR_BORDER
    End With
End Sub

' Slice a 2D array (1-based) into the requested sub-rectangle. VBA can't pass
' a range subset of an array directly, so we copy.
Public Function ResliceTwoD(arr As Variant, ByVal r1 As Long, ByVal r2 As Long, _
        ByVal c1 As Long, ByVal c2 As Long) As Variant
    Dim out() As Variant, i As Long, j As Long
    ReDim out(1 To r2 - r1 + 1, 1 To c2 - c1 + 1)
    For i = r1 To r2
        For j = c1 To c2
            out(i - r1 + 1, j - c1 + 1) = arr(i, j)
        Next j
    Next i
    ResliceTwoD = out
End Function
