Attribute VB_Name = "M01_Builder"
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
Option Explicit
Option Compare Text

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
    M02_DataEngine.ImportCsv
    M02_DataEngine.CleanData
    M02_DataEngine.BuildDataModel

    ' 4. Compute KPIs and AI-style analytics
    M03_KpiEngine.ComputeAllKpis
    M03_KpiEngine.RunAiAnalytics

    ' 5. Build all dashboard surfaces with premium UI
    M04_DashboardUI.BuildDashboardMain
    M06_Modules.BuildExecutiveView
    M06_Modules.BuildSlaIntelligence
    M06_Modules.BuildIncidentAnalytics
    M06_Modules.BuildAgentAnalytics
    M06_Modules.BuildApplicationAnalytics
    M06_Modules.BuildRcaAnalytics
    M06_Modules.BuildForecastAnalytics
    M06_Modules.BuildRiskAnalytics

    ' 6. Settings, logs and hidden config
    M07_Interaction.BuildSettingsSheet
    M07_Interaction.BuildLogsSheet
    M07_Interaction.BuildHiddenConfig

    ' 7. Apply navigation/filters/security to every analytics sheet
    M07_Interaction.WireNavigationEverywhere

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
Public Sub RefreshAll()
    On Error GoTo Fail
    PerfBegin
    LogInfo "REFRESH", "Refresh started"

    M02_DataEngine.ImportCsv
    M02_DataEngine.CleanData
    M02_DataEngine.BuildDataModel
    M03_KpiEngine.ComputeAllKpis
    M03_KpiEngine.RunAiAnalytics
    M04_DashboardUI.RefreshDashboardMain
    M06_Modules.RefreshAllModules

    LogInfo "REFRESH", "Done in " & Format(PerfEnd, "0.00") & "s"
    PerfRestore
    Application.StatusBar = "Refreshed at " & Format(Now, "hh:mm:ss")
    Exit Sub
Fail:
    PerfRestore
    LogError "REFRESH", Err.Number, Err.Description
End Sub

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
    Dim ws As Worksheet, keep As Worksheet
    Application.DisplayAlerts = False
    ' Add a temp sheet so we can delete everything else
    On Error Resume Next
    Set keep = ThisWorkbook.Worksheets.Add
    keep.Name = "__tmp_" & Format(Now, "hhnnss")
    On Error GoTo 0
    For Each ws In ThisWorkbook.Worksheets
        If Not ws Is keep Then ws.Delete
    Next ws
    Application.DisplayAlerts = True
End Sub

Private Sub ProvisionSheets()
    Dim names As Variant, i As Long, ws As Worksheet
    names = Array(SHT_HIDDEN, SHT_LOGS, SHT_SETTINGS, _
                  SHT_RAW, SHT_CLEAN, SHT_MODEL, SHT_KPI, _
                  SHT_RISK, SHT_FORECAST, SHT_RCA, SHT_APP, SHT_AGENT, _
                  SHT_INC, SHT_SLA, SHT_EXEC, SHT_DASH)
    ' Create in reverse so DASHBOARD ends up leftmost-visible
    For i = LBound(names) To UBound(names)
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = CStr(names(i))
        ws.Tab.Color = CLR_PANEL
        ws.DisplayPageBreaks = False
    Next i
    ' Drop the temp sheet
    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        If Left$(ThisWorkbook.Worksheets(i).Name, 6) = "__tmp_" Then ThisWorkbook.Worksheets(i).Delete
    Next i
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
