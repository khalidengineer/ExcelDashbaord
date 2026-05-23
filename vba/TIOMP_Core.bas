Attribute VB_Name = "TIOMP_Core"
Option Explicit
Option Compare Text

' === TIOMP_Core - Part of TIOMP Dashboard ===
' Imports: M01_Builder, M02_DataEngine, M03_KpiEngine


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
' Bullet-proof CSV locator. Searches every plausible location, then falls back
' to a file picker. Uses Application.GetOpenFilename (always available, no
' Microsoft Office Object Library reference required - safer than FileDialog).
'
' Priority order (first match wins):
'   1. Cached path stored on Hidden_Config sheet (after first successful find)
'   2. Workbook folder
'   3. Workbook folder \ data \
'   4. Workbook folder \ ExcelDashbaord \    (when extracted from GitHub ZIP)
'   5. Parent folder of workbook
'   6. Current working directory (CurDir)
'   7. User's Documents folder
'   8. User's Desktop folder
'   9. User's Downloads folder
'  10. File picker prompt (last resort)
Public Function ResolveCsvPath() As String
    On Error Resume Next   ' bullet-proof - swallow any path errors
    Dim sep As String: sep = Application.PathSeparator
    Dim base As String, p As String
    Dim profile As String: profile = Environ$("USERPROFILE")

    ' 0. Try cached path from Hidden_Config first
    Dim cached As String: cached = GetCachedCsvPath()
    If Len(cached) > 0 Then
        If Dir(cached) <> "" Then
            ResolveCsvPath = cached
            Exit Function
        End If
    End If

    ' 1-5. Try workbook folder + variants
    base = ThisWorkbook.Path
    If Len(base) > 0 Then
        p = base & sep & CSV_FILE_NAME
        If Dir(p) <> "" Then ResolveCsvPath = p: SaveCachedCsvPath p: Exit Function

        p = base & sep & "data" & sep & CSV_FILE_NAME
        If Dir(p) <> "" Then ResolveCsvPath = p: SaveCachedCsvPath p: Exit Function

        p = base & sep & "ExcelDashbaord" & sep & CSV_FILE_NAME
        If Dir(p) <> "" Then ResolveCsvPath = p: SaveCachedCsvPath p: Exit Function

        ' Parent folder - resolve ".." manually for old Excel safety
        Dim parent As String
        parent = ParentFolder(base)
        If Len(parent) > 0 Then
            p = parent & sep & CSV_FILE_NAME
            If Dir(p) <> "" Then ResolveCsvPath = p: SaveCachedCsvPath p: Exit Function
        End If
    End If

    ' 6. CurDir
    p = CurDir & sep & CSV_FILE_NAME
    If Dir(p) <> "" Then ResolveCsvPath = p: SaveCachedCsvPath p: Exit Function

    ' 7-9. User profile common folders
    If Len(profile) > 0 Then
        p = profile & sep & "Documents" & sep & CSV_FILE_NAME
        If Dir(p) <> "" Then ResolveCsvPath = p: SaveCachedCsvPath p: Exit Function

        p = profile & sep & "Desktop" & sep & CSV_FILE_NAME
        If Dir(p) <> "" Then ResolveCsvPath = p: SaveCachedCsvPath p: Exit Function

        p = profile & sep & "Downloads" & sep & CSV_FILE_NAME
        If Dir(p) <> "" Then ResolveCsvPath = p: SaveCachedCsvPath p: Exit Function
    End If

    ' 10. Last resort: file picker (uses GetOpenFilename - always available)
    Dim picked As Variant
    picked = Application.GetOpenFilename( _
        FileFilter:="CSV files (*.csv),*.csv,All files (*.*),*.*", _
        Title:="Select Apps.csv  -  not found automatically")
    If VarType(picked) = vbString Then
        If Len(CStr(picked)) > 0 Then
            ResolveCsvPath = CStr(picked)
            SaveCachedCsvPath CStr(picked)
            Exit Function
        End If
    End If

    ' Cancelled - return empty string. ImportCsv will raise a clear error.
    ResolveCsvPath = ""
End Function

' Resolve parent folder of a path, handling trailing separators.
Private Function ParentFolder(ByVal pth As String) As String
    On Error Resume Next
    Dim sep As String: sep = Application.PathSeparator
    Dim p As String: p = pth
    If Right$(p, 1) = sep Then p = Left$(p, Len(p) - 1)
    Dim pos As Long: pos = InStrRev(p, sep)
    If pos > 0 Then ParentFolder = Left$(p, pos - 1)
End Function

' Read cached CSV path from Hidden_Config sheet, cell B11
Private Function GetCachedCsvPath() As String
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_HIDDEN)
    If ws Is Nothing Then Exit Function
    GetCachedCsvPath = CStr(ws.Range("B11").Value)
End Function

' Persist successful CSV path so we don't re-search on every refresh
Private Sub SaveCachedCsvPath(ByVal pth As String)
    On Error Resume Next
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SHT_HIDDEN)
    If ws Is Nothing Then Exit Sub
    ws.Range("A11").Value = "CachedCsvPath"
    ws.Range("B11").Value = pth
End Sub


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
