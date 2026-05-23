Attribute VB_Name = "M07_Interaction"
'==============================================================================
' MODULE      : M07_Interaction
' DESCRIPTION : Interaction layer.  Navigation rail, command buttons, smart
'               search, drill-through, theme switching, PDF export, email
'               distribution, settings/logs scaffolding and login auth.
'==============================================================================
Option Explicit
Option Compare Text

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
    M04_DashboardUI.DeleteShapesByPrefix ws, "nav_"

    ' Side rail panel
    Dim panel As Shape
    Set panel = M04_DashboardUI.AddShape(ws, msoShapeRectangle, "nav_panel", _
        ws.Range("A1").Left, ws.Range("A1").Top + HEADER_HEIGHT + 4, _
        NAV_WIDTH - 6, 480)
    M04_DashboardUI.StyleShape panel, CLR_PANEL, 0
    M04_DashboardUI.AddShadow panel

    ' Section title
    Dim ttl As Shape
    Set ttl = M04_DashboardUI.AddShape(ws, msoShapeRectangle, "nav_ttl", _
        ws.Range("A1").Left + 8, ws.Range("A1").Top + HEADER_HEIGHT + 12, NAV_WIDTH - 22, 22)
    M04_DashboardUI.StyleShape ttl, CLR_PANEL, 0
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
    Set card = M04_DashboardUI.AddShape(ws, msoShapeRoundedRectangle, name, lft, tp, w, h)
    card.Adjustments.Item(1) = 0.3
    If isActive Then
        M04_DashboardUI.StyleShape card, CLR_ACCENT, 0
    Else
        M04_DashboardUI.StyleShape card, CLR_PANEL_HI, 0
    End If

    card.TextFrame2.TextRange.Text = "  " & glyph & "   " & label
    With card.TextFrame2.TextRange.Font
        .Size = 10: .Bold = msoTrue: .Name = "Segoe UI"
        If isActive Then .Fill.ForeColor.RGB = CLR_BG Else .Fill.ForeColor.RGB = CLR_TEXT
    End With
    card.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
    card.TextFrame2.VerticalAnchor = msoAnchorMiddle
    card.OnAction = "'M07_Interaction.NavigateTo """ & targetSheet & """'"
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
    Set btn = M04_DashboardUI.AddShape(ws, msoShapeRoundedRectangle, name, lft, tp, w, h)
    btn.Adjustments.Item(1) = 0.3
    M04_DashboardUI.StyleShape btn, fillCol, 0
    M04_DashboardUI.AddShadow btn
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
    Dim arr As Variant: arr = M03_KpiEngine.ReadCleanArrayPub
    If IsEmpty(arr) Then Exit Sub

    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Search_Results")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.name = "Search_Results"
    End If
    M04_DashboardUI.ResetSheetForDashboard ws
    M04_DashboardUI.PaintCanvas ws
    M04_DashboardUI.BuildHeaderBar ws, "SEARCH RESULTS  -  """ & term & """", _
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
            ws.Cells(outRow, col + 1).Interior.Color = M05_Charts.SeverityColor(CStr(arr(r, CD_NORM_SEVERITY)))
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
    Set pill = M04_DashboardUI.AddShape(ws, msoShapeRoundedRectangle, "srch_count", _
        ws.Cells(4, 2).Left, ws.Cells(4, 2).Top, 220, 22)
    M04_DashboardUI.StyleShape pill, CLR_PANEL_HI, 0
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
    h = h & MisRow("Total Tickets", M03_KpiEngine.FmtKpi("TotalTickets", "0"))
    h = h & MisRow("Open Tickets", M03_KpiEngine.FmtKpi("OpenTickets", "0"))
    h = h & MisRow("Critical Open", M03_KpiEngine.FmtKpi("CriticalOpen", "0"))
    h = h & MisRow("SLA Compliance", M03_KpiEngine.FmtKpi("SLA_Compliance", "0.0%"))
    h = h & MisRow("MTTR (h)", M03_KpiEngine.FmtKpi("MTTR_Hours", "0.00"))
    h = h & MisRow("Backlog Index", M03_KpiEngine.FmtKpi("Backlog_Index", "0.0%"))
    h = h & MisRow("Risk Score", M03_KpiEngine.FmtKpi("Risk_Score", "0.0"))
    h = h & MisRow("Ops Health", M03_KpiEngine.FmtKpi("OpsHealth_Score", "0.0"))
    h = h & MisRow("CSAT Index", M03_KpiEngine.FmtKpi("Customer_Satisfaction", "0.0"))
    h = h & MisRow("Trend WoW %", M03_KpiEngine.FmtKpi("Trend_Velocity", "0.0"))
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
    M04_DashboardUI.ResetSheetForDashboard ws
    M04_DashboardUI.PaintCanvas ws
    M04_DashboardUI.BuildHeaderBar ws, "SETTINGS  &  ADMINISTRATION", _
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
        Array(Array("Total Tickets", M03_KpiEngine.GetKpi("TotalTickets")), _
              Array("Open", M03_KpiEngine.GetKpi("OpenTickets")), _
              Array("Closed", M03_KpiEngine.GetKpi("ClosedTickets")), _
              Array("Last Refresh", Format(M03_KpiEngine.GetKpi("LastRefreshed"), "yyyy-mm-dd hh:mm")), _
              Array("Source", "Apps.csv"))

    ' Action buttons
    Dim btnLeft As Double, btnTop As Double
    btnLeft = ws.Cells(20, 2).Left
    btnTop = ws.Cells(20, 2).Top
    AddCommandButton ws, "set_btn_refresh", btnLeft, btnTop, 150, 32, _
        ChrW(8635) & "  REFRESH ALL", "M01_Builder.RefreshAll", CLR_ACCENT, CLR_BG
    AddCommandButton ws, "set_btn_rebuild", btnLeft + 160, btnTop, 150, 32, _
        ChrW(9881) & "  REBUILD ALL", "M01_Builder.BuildEnterpriseDashboard", CLR_PURPLE, vbWhite
    AddCommandButton ws, "set_btn_pdf", btnLeft + 320, btnTop, 150, 32, _
        ChrW(8595) & "  EXPORT PDF", "M07_Interaction.ExportDashboardPdf", CLR_ACCENT2, CLR_BG
    AddCommandButton ws, "set_btn_email", btnLeft + 480, btnTop, 150, 32, _
        ChrW(9993) & "  EMAIL MIS", "M07_Interaction.EmailMisReport", CLR_GOOD, vbWhite
    AddCommandButton ws, "set_btn_login", btnLeft + 640, btnTop, 150, 32, _
        ChrW(128272) & "  SIGN IN", "M07_Interaction.PromptLogin", CLR_PANEL_HI, CLR_TEXT
    AddCommandButton ws, "set_btn_backup", btnLeft, btnTop + 40, 150, 32, _
        ChrW(128190) & "  BACKUP", "M07_Interaction.BackupWorkbook", CLR_PANEL_HI, CLR_TEXT
    AddCommandButton ws, "set_btn_archive", btnLeft + 160, btnTop + 40, 150, 32, _
        ChrW(128193) & "  ARCHIVE CLOSED", "M07_Interaction.ArchiveClosed", CLR_PANEL_HI, CLR_TEXT
    AddCommandButton ws, "set_btn_clearlog", btnLeft + 320, btnTop + 40, 150, 32, _
        ChrW(128465) & "  CLEAR LOGS", "M07_Interaction.ClearLogs", CLR_BAD, vbWhite
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
    Dim arr As Variant: arr = M03_KpiEngine.ReadCleanArrayPub
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
