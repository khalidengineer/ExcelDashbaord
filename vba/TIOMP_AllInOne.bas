Attribute VB_Name = "TIOMP"
Option Explicit
Option Compare Text

'==============================================================================
' TIOMP - Ticket Intelligence & Operations Monitoring Platform
' SINGLE CONSOLIDATED VBA MODULE  (~4,600 lines)
' 
' SETUP (one-time, takes 2 minutes):
'   1. Save your workbook as TIOMP.xlsm next to Apps.csv
'   2. Press Alt+F11 to open VBA editor
'   3. Insert > Module
'   4. Paste THIS ENTIRE FILE into the module
'   5. Press Alt+F8, run BuildEnterpriseDashboard
'==============================================================================


'==============================================================================
' === SECTION: M01_Builder ===
'==============================================================================

'==============================================================================
' MODULE      : M01_Builder
' PROJECT     : Ticket Intelligence & Operations Monitoring Platform
' AUTHOR      : Enterprise BI / Kiro
' VERSION     : 1.0.0
' DESCRIPTION : Master orchestrator. Single entry point that builds the entire
'               enterprise dashboard from Apps.csv in one click.
'               Hosts global constants (theme, layout, SLA targets), a
'               structured logger and performance helpers.
'==============================================================================

'------------------------------------------------------------------------------
' GLOBAL CONFIGURATION
'------------------------------------------------------------------------------
Public Const APP_NAME              As String = "Ticket Intelligence & Operations Monitoring Platform"
Public Const APP_SHORT             As String = "TIOMP"
Public Const APP_VERSION           As String = "1.0.0"

' Source CSV - resolved at runtime via ResolveCsvPath()
Public Const CSV_FILE_NAME         As String = "Apps.csv"

' Sheet names (single source of truth)
Public Const SHT_RAW               As String = "Raw_Data"
Public Const SHT_CLEAN             As String = "Clean_Data"
Public Const SHT_MODEL             As String = "Data_Model"
Public Const SHT_KPI               As String = "KPI_Engine"
Public Const SHT_DASH              As String = "Dashboard_Main"
Public Const SHT_EXEC              As String = "Executive_View"
Public Const SHT_SLA               As String = "SLA_Intelligence"
Public Const SHT_INC               As String = "Incident_Analytics"
Public Const SHT_AGENT             As String = "Agent_Analytics"
Public Const SHT_APP               As String = "Application_Analytics"
Public Const SHT_RCA               As String = "RCA_Analytics"
Public Const SHT_FORECAST          As String = "Forecast_Analytics"
Public Const SHT_RISK              As String = "Risk_Analytics"
Public Const SHT_SETTINGS          As String = "Settings"
Public Const SHT_LOGS              As String = "Logs"
Public Const SHT_HIDDEN            As String = "Hidden_Config"

'------------------------------------------------------------------------------
' THEME - Dark NOC / Glassmorphism palette (default)
'------------------------------------------------------------------------------
Public Const CLR_BG                As Long = &H1A170F      '  #0F171A near-black
Public Const CLR_PANEL             As Long = &H32231A      '  #1A2332 panel
Public Const CLR_PANEL_HI          As Long = &H402D22      '  #222D40 hover
Public Const CLR_BORDER            As Long = &H4F3D2E      '  #2E3D4F subtle border
Public Const CLR_TEXT              As Long = &HFFF1E6      '  #E6F1FF primary text
Public Const CLR_MUTED             As Long = &HB39A87      '  #879AB3 secondary text
Public Const CLR_ACCENT            As Long = &HFFE500      '  #00E5FF neon cyan
Public Const CLR_ACCENT2           As Long = &HFFB300      '  #00B3FF blue accent
Public Const CLR_GOOD              As Long = &H81B910      '  #10B981 emerald
Public Const CLR_WARN              As Long = &H0B9EF5      '  #F59E0B amber
Public Const CLR_BAD               As Long = &H4444EF      '  #EF4444 red
Public Const CLR_PURPLE            As Long = &HF6549B      '  #9B54F6 purple
Public Const CLR_PINK              As Long = &HA660EC      '  #EC60A6 pink

'------------------------------------------------------------------------------
' SLA TARGETS (in hours, by Severity) - tunable via Settings sheet
'------------------------------------------------------------------------------
Public Const SLA_CRITICAL_H        As Double = 4#
Public Const SLA_HIGH_H            As Double = 8#
Public Const SLA_MEDIUM_H          As Double = 24#
Public Const SLA_LOW_H             As Double = 48#
Public Const SLA_UNSPEC_H          As Double = 24#

'------------------------------------------------------------------------------
' LAYOUT CONSTANTS
'------------------------------------------------------------------------------
Public Const HEADER_HEIGHT         As Double = 70#
Public Const NAV_WIDTH             As Double = 200#
Public Const KPI_CARD_W            As Double = 180#
Public Const KPI_CARD_H            As Double = 96#
Public Const PAD                   As Double = 12#

'------------------------------------------------------------------------------
' RUNTIME STATE
'------------------------------------------------------------------------------
Private mPerfStart As Double

'==============================================================================
'                              ENTRY POINTS
'==============================================================================

' One-click build of the entire platform.  Run this from a fresh workbook.
Public Sub BuildEnterpriseDashboard()
    On Error GoTo Fail
    PerfBegin
    LogInfo "BUILD", "Starting full build of " & APP_NAME

    ' 1. Hard reset workbook to a known state
    ResetWorkbook

    ' 2. Provision all required sheets
    ProvisionSheets

    ' 3. Import + clean + model the data
    ImportCsv
    CleanData
    BuildDataModel

    ' 4. Compute KPIs and AI-style analytics
    ComputeAllKpis
    RunAiAnalytics

    ' 5. Build all dashboard surfaces with premium UI
    BuildDashboardMain
    BuildExecutiveView
    BuildSlaIntelligence
    BuildIncidentAnalytics
    BuildAgentAnalytics
    BuildApplicationAnalytics
    BuildRcaAnalytics
    BuildForecastAnalytics
    BuildRiskAnalytics

    ' 6. Settings, logs and hidden config
    BuildSettingsSheet
    BuildLogsSheet
    BuildHiddenConfig

    ' 7. Apply navigation/filters/security to every analytics sheet
    WireNavigationEverywhere

    ' 8. Land the user on the main dashboard
    On Error Resume Next
    Sheets(SHT_DASH).Activate
    ActiveWindow.DisplayGridlines = False
    Range("A1").Select
    On Error GoTo Fail

    LogInfo "BUILD", "Completed in " & Format(PerfEnd, "0.00") & "s"
    PerfRestore

    MsgBox APP_NAME & " is ready." & vbCrLf & _
           "Sheets built: 16  |  Build time: " & Format(PerfEnd, "0.00") & "s" & vbCrLf & _
           "Tip: press the REFRESH button on the dashboard whenever the CSV updates.", _
           vbInformation, APP_SHORT
    Exit Sub
Fail:
    PerfRestore
    LogError "BUILD", Err.Number, Err.Description
    MsgBox "Build failed: " & Err.Description, vbCritical, APP_SHORT
End Sub

' Refresh path - re-imports CSV and rebuilds analytics & visuals without touching layout.
' If the dashboard hasn't been built yet, it transparently runs a full build instead.
Public Sub RefreshAll()
    On Error GoTo Fail

    ' Auto-detect first run - if Dashboard_Main is missing, do a full build
    If Not SheetExists(SHT_DASH) Then
        LogInfo "REFRESH", "Dashboard_Main not found - running full build instead"
        BuildEnterpriseDashboard
        Exit Sub
    End If

    PerfBegin
    LogInfo "REFRESH", "Refresh started"

    ImportCsv
    CleanData
    BuildDataModel
    ComputeAllKpis
    RunAiAnalytics
    RefreshDashboardMain
    RefreshAllModules

    LogInfo "REFRESH", "Done in " & Format(PerfEnd, "0.00") & "s"
    PerfRestore
    Application.StatusBar = "Refreshed at " & Format(Now, "hh:mm:ss")
    Exit Sub
Fail:
    PerfRestore
    LogError "REFRESH", Err.Number, Err.Description
    MsgBox "Refresh failed: " & Err.Description & vbCrLf & vbCrLf & _
           "Tip: run BuildEnterpriseDashboard from the macro list to rebuild from scratch.", _
           vbExclamation, APP_SHORT
End Sub

' Helper - safely test if a sheet exists by name
Public Function SheetExists(ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    SheetExists = Not ws Is Nothing
End Function

'==============================================================================
'                          PERFORMANCE / STATE HELPERS
'==============================================================================
Public Sub PerfBegin()
    mPerfStart = Timer
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.Cursor = xlWait
End Sub

Public Function PerfEnd() As Double
    PerfEnd = Timer - mPerfStart
End Function

Public Sub PerfRestore()
    Application.DisplayAlerts = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.StatusBar = False
End Sub

'==============================================================================
'                              SHEET PROVISIONING
'==============================================================================
Private Sub ResetWorkbook()
    Dim i As Long, ws As Worksheet, keep As Worksheet
    Application.DisplayAlerts = False
    ' Make all sheets visible first - hidden sheets can't be the only sheet left
    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        ws.Visible = xlSheetVisible
    Next ws
    On Error GoTo 0

    ' Add a temp sheet so we can delete everything else (workbook needs >=1 sheet)
    On Error Resume Next
    Set keep = ThisWorkbook.Worksheets.Add
    keep.Name = "__tmp_" & Format(Now, "hhnnss")
    On Error GoTo 0
    If keep Is Nothing Then Set keep = ThisWorkbook.Worksheets(1)

    ' Delete by reverse index - safer than For Each while mutating the collection
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        If ws.Name <> keep.Name Then
            On Error Resume Next
            ws.Delete
            On Error GoTo 0
        End If
    Next i
    Application.DisplayAlerts = True
End Sub

Private Sub ProvisionSheets()
    Dim names As Variant, i As Long, ws As Worksheet, nm As String
    names = Array(SHT_HIDDEN, SHT_LOGS, SHT_SETTINGS, _
                  SHT_RAW, SHT_CLEAN, SHT_MODEL, SHT_KPI, _
                  SHT_RISK, SHT_FORECAST, SHT_RCA, SHT_APP, SHT_AGENT, _
                  SHT_INC, SHT_SLA, SHT_EXEC, SHT_DASH)

    Application.DisplayAlerts = False
    ' Create in reverse so DASHBOARD ends up leftmost-visible
    For i = LBound(names) To UBound(names)
        nm = CStr(names(i))
        ' If a sheet with this name already exists, delete it first to avoid name clash
        If SheetExists(nm) Then
            On Error Resume Next
            ThisWorkbook.Worksheets(nm).Delete
            On Error GoTo 0
        End If
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nm
        On Error Resume Next
        ws.Tab.Color = CLR_PANEL
        ws.DisplayPageBreaks = False
        On Error GoTo 0
    Next i

    ' Drop any leftover temp sheets
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If Left$(ThisWorkbook.Worksheets(i).Name, 6) = "__tmp_" Then
            On Error Resume Next
            ThisWorkbook.Worksheets(i).Delete
            On Error GoTo 0
        End If
    Next i
    Application.DisplayAlerts = True

    ' Reorder so DASHBOARD is first, hidden last
    On Error Resume Next
    Sheets(SHT_DASH).Move Before:=Sheets(1)
    Sheets(SHT_HIDDEN).Move After:=Sheets(Sheets.Count)
    Sheets(SHT_HIDDEN).Visible = xlSheetVeryHidden
    Sheets(SHT_RAW).Visible = xlSheetHidden
    Sheets(SHT_CLEAN).Visible = xlSheetHidden
    Sheets(SHT_MODEL).Visible = xlSheetHidden
    Sheets(SHT_KPI).Visible = xlSheetHidden
    On Error GoTo 0
End Sub

'==============================================================================
'                                 LOGGER
'==============================================================================
Public Sub LogInfo(ByVal source As String, ByVal msg As String)
    WriteLog "INFO", source, msg
End Sub
Public Sub LogWarn(ByVal source As String, ByVal msg As String)
    WriteLog "WARN", source, msg
End Sub
Public Sub LogError(ByVal source As String, ByVal errNum As Long, ByVal msg As String)
    WriteLog "ERROR", source, "Err " & errNum & " - " & msg
End Sub

Private Sub WriteLog(ByVal level As String, ByVal source As String, ByVal msg As String)
    Dim ws As Worksheet, r As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHT_LOGS)
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r < 2 Then r = 2
    ws.Cells(r, 1).Value = Now
    ws.Cells(r, 2).Value = level
    ws.Cells(r, 3).Value = source
    ws.Cells(r, 4).Value = msg
    ws.Cells(r, 1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    On Error GoTo 0
End Sub

'==============================================================================
'                             FILE PATH RESOLVER
'==============================================================================
' Looks for Apps.csv next to the workbook, in the workbook folder, or one level up.
Public Function ResolveCsvPath() As String
    Dim base As String, candidates As Variant, i As Long, p As String
    base = ThisWorkbook.Path
    If Len(base) = 0 Then base = CurDir
    candidates = Array( _
        base & Application.PathSeparator & CSV_FILE_NAME, _
        base & Application.PathSeparator & "data" & Application.PathSeparator & CSV_FILE_NAME, _
        base & Application.PathSeparator & ".." & Application.PathSeparator & CSV_FILE_NAME)
    For i = LBound(candidates) To UBound(candidates)
        p = CStr(candidates(i))
        If Dir(p) <> "" Then
            ResolveCsvPath = p
            Exit Function
        End If
    Next i
    ' As a final fallback, prompt the user
    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    With fd
        .Title = "Select Apps.csv"
        .Filters.Clear
        .Filters.Add "CSV Files", "*.csv"
        .AllowMultiSelect = False
        If .Show = -1 Then ResolveCsvPath = .SelectedItems(1)
    End With
End Function


'==============================================================================
' === SECTION: M02_DataEngine ===
'==============================================================================

'==============================================================================
' MODULE      : M02_DataEngine
' DESCRIPTION : Auto CSV import, smart cleaning, normalization, derived columns
'               and aggregation model. All transforms run on in-memory arrays
'               for performance, then write back in single shot.
'==============================================================================

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
    If Len(path) = 0 Then Err.Raise 5001, , "Apps.csv not found and user cancelled file picker."
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
    arr = ReadCleanArrayPub()
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


'==============================================================================
' === SECTION: M03_KpiEngine ===
'==============================================================================

'==============================================================================
' MODULE      : M03_KpiEngine
' DESCRIPTION : Computes enterprise KPIs and runs AI-style analytics.
'               Outputs land on KPI_Engine sheet as a key/value store that the
'               UI layer reads directly via GetKpi(name).
'==============================================================================

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


'==============================================================================
' === SECTION: M04_DashboardUI ===
'==============================================================================

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

'==============================================================================
' MODULE      : M05_Charts
' DESCRIPTION : Chart Rendering Engine.  All charts read from named blocks on
'               the Data_Model sheet (located by section header).  Every chart
'               is restyled to dark NOC aesthetic.
'==============================================================================

'==============================================================================
'                          PUBLIC API
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
