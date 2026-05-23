VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'==============================================================================
' MODULE      : ThisWorkbook
' DESCRIPTION : Workbook-level event handlers.  Auto-prompts user to refresh
'               the dashboard on open and registers a clean shut-down.
'==============================================================================
Option Explicit

Private Sub Workbook_Open()
    On Error Resume Next
    Application.Cursor = xlDefault
    ' If the dashboard sheet exists, just navigate.  If first time (only one sheet
    ' visible), offer to build it.
    Dim hasDash As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.name = "Dashboard_Main" Then hasDash = True: Exit For
    Next ws

    If Not hasDash Then
        Dim ans As VbMsgBoxResult
        ans = MsgBox("Welcome to Ticket Intelligence & Operations Monitoring Platform." & vbCrLf & vbCrLf & _
                     "No dashboard exists yet. Would you like to build it now from Apps.csv?", _
                     vbYesNo + vbInformation, "TIOMP - First Run")
        If ans = vbYes Then M01_Builder.BuildEnterpriseDashboard
    Else
        ' Refresh on open if user opted-in
        Dim ans2 As VbMsgBoxResult
        ans2 = MsgBox("Refresh dashboard with the latest Apps.csv data?", _
                      vbYesNo + vbQuestion + vbDefaultButton2, "TIOMP - Refresh")
        If ans2 = vbYes Then M01_Builder.RefreshAll
        ThisWorkbook.Worksheets("Dashboard_Main").Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
    End If
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    On Error Resume Next
    M01_Builder.LogInfo "WB", "Workbook closing"
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.Cursor = xlDefault
End Sub

Private Sub Workbook_SheetActivate(ByVal Sh As Object)
    On Error Resume Next
    ' Hide gridlines & headings on every analytics surface for a polished look
    Select Case Sh.name
        Case "Dashboard_Main", "Executive_View", "SLA_Intelligence", _
             "Incident_Analytics", "Agent_Analytics", "Application_Analytics", _
             "RCA_Analytics", "Forecast_Analytics", "Risk_Analytics", _
             "Settings", "Search_Results"
            ActiveWindow.DisplayGridlines = False
            ActiveWindow.DisplayHeadings = False
    End Select
End Sub
