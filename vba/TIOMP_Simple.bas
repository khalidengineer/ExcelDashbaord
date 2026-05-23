Attribute VB_Name = "TIOMP_Simple"
'==============================================================================
' TIOMP SIMPLE - Bulletproof Ticket Dashboard
' Works on Excel 2007+ - uses ONLY basic features. No fancy shapes/shadows.
'
' SETUP:
'   1. Save workbook as TIOMP.xlsm next to Apps.csv
'   2. Alt+F11, Insert > Module, paste this entire file
'   3. Alt+F8, run BuildDashboard
'==============================================================================
Option Explicit

' Sheet names
Public Const SH_DATA As String = "Data"
Public Const SH_KPI As String = "KPI"
Public Const SH_DASH As String = "Dashboard"

' Colors (decimal RGB - safe for all Excel versions)
Public Const CLR_HEADER As Long = 2304060        ' dark blue/black
Public Const CLR_ACCENT As Long = 16746496       ' blue
Public Const CLR_GOOD As Long = 5287936          ' green
Public Const CLR_WARN As Long = 39423            ' amber
Public Const CLR_BAD As Long = 4474111           ' red
Public Const CLR_PANEL As Long = 3289650         ' gray
Public Const CLR_TEXT As Long = 16777215         ' white

'==============================================================================
' MAIN ENTRY POINT
'==============================================================================
Public Sub BuildDashboard()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    ' Step 1: Find and import CSV
    Dim csvPath As String
    csvPath = FindCsv()
    If Len(csvPath) = 0 Then
        MsgBox "Apps.csv not found. Please put it next to this workbook." & vbCrLf & vbCrLf & _
               "Expected location: " & ThisWorkbook.Path, vbExclamation, "TIOMP"
        GoTo CleanExit
    End If

    ' Step 2: Reset workbook
    ResetSheets

    ' Step 3: Import data
    ImportCsvData csvPath

    ' Step 4: Build KPIs
    BuildKpis

    ' Step 5: Build dashboard
    BuildDashboardSheet

    ' Step 6: Activate dashboard
    Sheets(SH_DASH).Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    Range("A1").Select
    On Error GoTo ErrHandler

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Dashboard ready! Tickets imported successfully.", vbInformation, "TIOMP"
    Exit Sub

ErrHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & _
           "On line: " & Erl, vbCritical, "TIOMP Build Failed"
    Exit Sub
CleanExit:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub

'==============================================================================
' FIND CSV - searches multiple folders
'==============================================================================
Public Function FindCsv() As String
    On Error Resume Next
    Dim sep As String: sep = Application.PathSeparator
    Dim base As String: base = ThisWorkbook.Path
    Dim profile As String: profile = Environ$("USERPROFILE")
    Dim p As String

    ' Try workbook folder
    If Len(base) > 0 Then
        p = base & sep & "Apps.csv"
        If Dir(p) <> "" Then FindCsv = p: Exit Function
    End If
    ' Try CurDir
    p = CurDir & sep & "Apps.csv"
    If Dir(p) <> "" Then FindCsv = p: Exit Function
    ' Try common user folders
    If Len(profile) > 0 Then
        p = profile & sep & "Documents" & sep & "Apps.csv"
        If Dir(p) <> "" Then FindCsv = p: Exit Function
        p = profile & sep & "Desktop" & sep & "Apps.csv"
        If Dir(p) <> "" Then FindCsv = p: Exit Function
        p = profile & sep & "Downloads" & sep & "Apps.csv"
        If Dir(p) <> "" Then FindCsv = p: Exit Function
    End If
    ' Last resort - file picker
    Dim picked As Variant
    picked = Application.GetOpenFilename( _
        FileFilter:="CSV files (*.csv),*.csv", _
        Title:="Select Apps.csv")
    If VarType(picked) = vbString Then FindCsv = CStr(picked)
End Function

'==============================================================================
' RESET SHEETS
'==============================================================================
Public Sub ResetSheets()
    Dim sheetNames As Variant
    sheetNames = Array(SH_DATA, SH_KPI, SH_DASH)
    Dim i As Long, ws As Worksheet, nm As String

    ' Make sure at least one sheet stays visible while we work
    On Error Resume Next
    Dim tmp As Worksheet
    Set tmp = Worksheets.Add
    tmp.Name = "_tmp_tiomp"
    On Error GoTo 0

    ' Delete our 3 target sheets if they exist
    For i = LBound(sheetNames) To UBound(sheetNames)
        nm = CStr(sheetNames(i))
        On Error Resume Next
        Sheets(nm).Delete
        On Error GoTo 0
    Next i

    ' Create fresh sheets
    For i = LBound(sheetNames) To UBound(sheetNames)
        nm = CStr(sheetNames(i))
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        ws.Name = nm
    Next i

    ' Clean up temp
    On Error Resume Next
    Sheets("_tmp_tiomp").Delete
    On Error GoTo 0

    ' Reorder - Dashboard first
    On Error Resume Next
    Sheets(SH_DASH).Move Before:=Sheets(1)
    On Error GoTo 0
End Sub

'==============================================================================
' IMPORT CSV - line by line, handles quoted commas
'==============================================================================
Public Sub ImportCsvData(ByVal path As String)
    Dim ws As Worksheet
    Set ws = Sheets(SH_DATA)
    ws.Cells.Clear

    Dim fnum As Integer
    fnum = FreeFile

    Dim allLines As Collection
    Set allLines = New Collection

    Dim lineText As String
    Open path For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, lineText
        ' Strip BOM
        If Len(lineText) > 0 Then
            If AscW(Left$(lineText, 1)) = 65279 Then lineText = Mid$(lineText, 2)
        End If
        allLines.Add lineText
    Loop
    Close #fnum

    If allLines.Count = 0 Then Exit Sub

    Dim numCols As Long: numCols = 24
    Dim outArr() As Variant
    ReDim outArr(1 To allLines.Count, 1 To numCols)

    Dim r As Long, c As Long
    Dim fields As Variant
    For r = 1 To allLines.Count
        fields = ParseCsvLineSimple(CStr(allLines(r)))
        For c = 1 To numCols
            If c <= UBound(fields) - LBound(fields) + 1 Then
                outArr(r, c) = fields(LBound(fields) + c - 1)
            End If
        Next c
    Next r

    ws.Range(ws.Cells(1, 1), ws.Cells(allLines.Count, numCols)).Value = outArr

    ' Style header row
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, numCols))
        .Interior.Color = CLR_HEADER
        .Font.Color = CLR_TEXT
        .Font.Bold = True
        .Font.Name = "Calibri"
        .Font.Size = 10
    End With
    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 9
    ws.Columns.AutoFit
End Sub

'==============================================================================
' SIMPLE CSV LINE PARSER - handles quoted commas
'==============================================================================
Public Function ParseCsvLineSimple(ByVal s As String) As Variant
    Dim out() As String
    ReDim out(0 To 0)
    Dim n As Long: n = 0
    Dim cur As String, ch As String
    Dim i As Long, inQuote As Boolean
    cur = ""
    inQuote = False

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = """" Then
            If inQuote And i < Len(s) Then
                If Mid$(s, i + 1, 1) = """" Then
                    cur = cur & """"
                    i = i + 1
                Else
                    inQuote = False
                End If
            Else
                inQuote = Not inQuote
            End If
        ElseIf ch = "," And Not inQuote Then
            ReDim Preserve out(0 To n)
            out(n) = cur
            n = n + 1
            cur = ""
        Else
            cur = cur & ch
        End If
    Next i
    ReDim Preserve out(0 To n)
    out(n) = cur
    ParseCsvLineSimple = out
End Function

'==============================================================================
' BUILD KPIs - calculate totals into KPI sheet
'==============================================================================
Public Sub BuildKpis()
    Dim wsData As Worksheet, wsKpi As Worksheet
    Set wsData = Sheets(SH_DATA)
    Set wsKpi = Sheets(SH_KPI)
    wsKpi.Cells.Clear

    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, 9).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    Dim totalTickets As Long, openCount As Long, closedCount As Long
    Dim critCount As Long, highCount As Long, medCount As Long, lowCount As Long
    Dim itCount As Long, rmCount As Long, otherCount As Long
    Dim openCritical As Long
    Dim r As Long
    Dim sev As String, stat As String, itype As String

    ' Severity dictionaries
    Dim sevDict As Object: Set sevDict = CreateObject("Scripting.Dictionary")
    Dim statDict As Object: Set statDict = CreateObject("Scripting.Dictionary")
    Dim typeDict As Object: Set typeDict = CreateObject("Scripting.Dictionary")
    Dim locDict As Object: Set locDict = CreateObject("Scripting.Dictionary")
    Dim resDict As Object: Set resDict = CreateObject("Scripting.Dictionary")

    For r = 2 To lastRow
        totalTickets = totalTickets + 1
        sev = NormSev(CStr(wsData.Cells(r, 12).Value))
        stat = NormStat(CStr(wsData.Cells(r, 13).Value))
        itype = NormType(CStr(wsData.Cells(r, 11).Value))

        ' Counts
        If stat = "Open" Then openCount = openCount + 1
        If stat = "Closed" Then closedCount = closedCount + 1
        If sev = "Critical" Then critCount = critCount + 1
        If sev = "High" Then highCount = highCount + 1
        If sev = "Medium" Then medCount = medCount + 1
        If sev = "Low" Then lowCount = lowCount + 1
        If sev = "Critical" And stat = "Open" Then openCritical = openCritical + 1

        ' Dictionaries
        AddCount sevDict, sev
        AddCount statDict, stat
        AddCount typeDict, itype
        AddCount locDict, CStr(wsData.Cells(r, 7).Value)
        Dim resolver As String: resolver = Trim$(CStr(wsData.Cells(r, 16).Value))
        If Len(resolver) > 0 Then AddCount resDict, resolver
    Next r

    ' Write headline KPIs (col A, B)
    wsKpi.Range("A1").Value = "KPI"
    wsKpi.Range("B1").Value = "Value"
    wsKpi.Range("A1:B1").Interior.Color = CLR_HEADER
    wsKpi.Range("A1:B1").Font.Color = CLR_TEXT
    wsKpi.Range("A1:B1").Font.Bold = True

    Dim row As Long: row = 2
    PutKpi wsKpi, row, "Total Tickets", totalTickets
    PutKpi wsKpi, row, "Open", openCount
    PutKpi wsKpi, row, "Closed", closedCount
    PutKpi wsKpi, row, "Critical", critCount
    PutKpi wsKpi, row, "Critical Open", openCritical
    PutKpi wsKpi, row, "High", highCount
    PutKpi wsKpi, row, "Medium", medCount
    PutKpi wsKpi, row, "Low", lowCount
    Dim closurePct As Double
    If totalTickets > 0 Then closurePct = closedCount / totalTickets Else closurePct = 0
    PutKpi wsKpi, row, "Closure Rate %", closurePct
    wsKpi.Cells(row - 1, 2).NumberFormat = "0.0%"

    ' Severity table - col D, E
    BuildAggTable wsKpi, 1, 4, "Severity", "Count", sevDict
    ' Status - col G, H
    BuildAggTable wsKpi, 1, 7, "Status", "Count", statDict
    ' Issue Type - col J, K
    BuildAggTable wsKpi, 1, 10, "Issue Type", "Count", typeDict
    ' Top Locations - col M, N
    BuildAggTable wsKpi, 1, 13, "Location", "Count", locDict, 10
    ' Top Resolvers - col P, Q
    BuildAggTable wsKpi, 1, 16, "Resolver", "Count", resDict, 10

    wsKpi.Columns.AutoFit
End Sub

Public Sub PutKpi(ws As Worksheet, ByRef row As Long, ByVal label As String, ByVal value As Variant)
    ws.Cells(row, 1).Value = label
    ws.Cells(row, 2).Value = value
    row = row + 1
End Sub

Public Sub AddCount(d As Object, ByVal key As String)
    If Len(Trim$(key)) = 0 Then key = "(blank)"
    If d.Exists(key) Then
        d(key) = d(key) + 1
    Else
        d(key) = 1
    End If
End Sub

Public Sub BuildAggTable(ws As Worksheet, ByVal startRow As Long, ByVal startCol As Long, _
        ByVal hdr1 As String, ByVal hdr2 As String, d As Object, Optional ByVal topN As Long = 99)
    ws.Cells(startRow, startCol).Value = hdr1
    ws.Cells(startRow, startCol + 1).Value = hdr2
    With ws.Range(ws.Cells(startRow, startCol), ws.Cells(startRow, startCol + 1))
        .Interior.Color = CLR_HEADER
        .Font.Color = CLR_TEXT
        .Font.Bold = True
    End With

    ' Sort keys by count desc
    Dim keys As Variant: keys = d.keys
    Dim n As Long: n = d.Count
    If n = 0 Then Exit Sub
    Dim i As Long, j As Long, t As Variant
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If d(keys(i)) < d(keys(j)) Then
                t = keys(i): keys(i) = keys(j): keys(j) = t
            End If
        Next j
    Next i

    Dim limit As Long
    limit = n - 1
    If limit > topN - 1 Then limit = topN - 1

    For i = 0 To limit
        ws.Cells(startRow + 1 + i, startCol).Value = keys(i)
        ws.Cells(startRow + 1 + i, startCol + 1).Value = d(keys(i))
    Next i
End Sub

'==============================================================================
' NORMALIZERS
'==============================================================================
Public Function NormSev(ByVal s As String) As String
    Dim t As String: t = LCase$(Trim$(s))
    Select Case t
        Case "critical": NormSev = "Critical"
        Case "high": NormSev = "High"
        Case "medium": NormSev = "Medium"
        Case "low": NormSev = "Low"
        Case "": NormSev = "Unspecified"
        Case Else: NormSev = s
    End Select
End Function

Public Function NormStat(ByVal s As String) As String
    Dim t As String: t = LCase$(Trim$(s))
    Select Case t
        Case "closed", "resolved", "done": NormStat = "Closed"
        Case Else: NormStat = "Open"
    End Select
End Function

Public Function NormType(ByVal s As String) As String
    Dim t As String: t = LCase$(Trim$(s))
    If t = "" Then
        NormType = "Uncategorized"
    ElseIf t = "it" Then
        NormType = "IT"
    ElseIf InStr(t, "repair") > 0 Or InStr(t, "maintenance") > 0 Then
        NormType = "Repair & Maintenance"
    Else
        NormType = s
    End If
End Function

'==============================================================================
' BUILD DASHBOARD SHEET - all using cells, no shapes
'==============================================================================
Public Sub BuildDashboardSheet()
    Dim ws As Worksheet
    Set ws = Sheets(SH_DASH)
    ws.Cells.Clear

    ' Delete any existing charts
    On Error Resume Next
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co
    On Error GoTo 0

    ' Set column widths
    Dim c As Long
    For c = 1 To 16
        ws.Columns(c).ColumnWidth = 14
    Next c
    ws.Cells.Font.Name = "Calibri"

    ' Title bar (row 1)
    ws.Range("A1:P1").Merge
    ws.Range("A1").Value = "TICKET INTELLIGENCE & OPERATIONS DASHBOARD"
    With ws.Range("A1")
        .Font.Size = 18
        .Font.Bold = True
        .Font.Color = CLR_TEXT
        .Interior.Color = CLR_HEADER
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    ws.Rows(1).RowHeight = 36

    ' Subtitle (row 2)
    ws.Range("A2:P2").Merge
    ws.Range("A2").Value = "Built " & Format(Now, "dd-mmm-yyyy hh:mm") & " | Source: Apps.csv"
    With ws.Range("A2")
        .Font.Size = 10
        .Font.Italic = True
        .Interior.Color = CLR_PANEL
        .Font.Color = CLR_TEXT
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(2).RowHeight = 22

    ' KPI Cards - 4 columns wide x 2 rows tall, starting row 4
    Dim wsKpi As Worksheet
    Set wsKpi = Sheets(SH_KPI)

    DrawKpiCard ws, "B4", "D6", "TOTAL TICKETS", wsKpi.Range("B2").Value, "0", CLR_ACCENT
    DrawKpiCard ws, "E4", "G6", "OPEN", wsKpi.Range("B3").Value, "0", CLR_WARN
    DrawKpiCard ws, "H4", "J6", "CLOSED", wsKpi.Range("B4").Value, "0", CLR_GOOD
    DrawKpiCard ws, "K4", "M6", "CRITICAL OPEN", wsKpi.Range("B6").Value, "0", CLR_BAD

    DrawKpiCard ws, "B8", "D10", "CRITICAL", wsKpi.Range("B5").Value, "0", CLR_BAD
    DrawKpiCard ws, "E8", "G10", "HIGH", wsKpi.Range("B7").Value, "0", CLR_WARN
    DrawKpiCard ws, "H8", "J10", "CLOSURE RATE", wsKpi.Range("B10").Value, "0.0%", CLR_GOOD
    DrawKpiCard ws, "K8", "M10", "MEDIUM+LOW", wsKpi.Range("B8").Value + wsKpi.Range("B9").Value, "0", CLR_ACCENT

    ' Section: Severity / Status / Issue Type tables
    DrawSection ws, 12, "BREAKDOWN BY SEVERITY, STATUS, AND ISSUE TYPE"
    CopyTable wsKpi, "D1:E10", ws, "B14"
    CopyTable wsKpi, "G1:H10", ws, "F14"
    CopyTable wsKpi, "J1:K10", ws, "J14"

    ' Section: Top Locations
    DrawSection ws, 24, "TOP REPORTING LOCATIONS"
    CopyTable wsKpi, "M1:N12", ws, "B26"

    ' Section: Top Resolvers
    DrawSection ws, 24, "TOP RESOLVERS", 9
    CopyTable wsKpi, "P1:Q12", ws, "I26"

    ' Charts
    BuildSeverityChart ws, "B40"
    BuildLocationChart ws, "I40"

    ' Add refresh button
    AddRefreshButton ws

    ws.Range("A1").Select
End Sub

Public Sub DrawKpiCard(ws As Worksheet, ByVal topLeft As String, ByVal botRight As String, _
        ByVal label As String, ByVal value As Variant, ByVal numFmt As String, ByVal accentColor As Long)
    Dim rng As Range
    Set rng = ws.Range(topLeft & ":" & botRight)
    rng.Merge
    With rng
        .Interior.Color = CLR_PANEL
        .Borders.LineStyle = xlContinuous
        .Borders.Color = accentColor
        .Borders.Weight = xlMedium
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Color = CLR_TEXT
        .Font.Bold = True
    End With

    ' Use line break - put value (big) and label (small)
    Dim displayVal As String
    If IsNumeric(value) Then
        displayVal = Format(value, numFmt)
    Else
        displayVal = CStr(value)
    End If
    rng.Value = displayVal & vbLf & label
    rng.Font.Size = 16
    rng.WrapText = True
End Sub

Public Sub DrawSection(ws As Worksheet, ByVal row As Long, ByVal title As String, _
        Optional ByVal startCol As Long = 2)
    Dim endCol As Long: endCol = 13
    If startCol > 2 Then endCol = startCol + 5
    ws.Cells(row, startCol).Value = title
    With ws.Range(ws.Cells(row, startCol), ws.Cells(row, endCol))
        .Merge
        .Font.Size = 12
        .Font.Bold = True
        .Font.Color = CLR_TEXT
        .Interior.Color = CLR_HEADER
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .IndentLevel = 1
    End With
    ws.Rows(row).RowHeight = 24
End Sub

Public Sub CopyTable(srcWs As Worksheet, ByVal srcAddr As String, dstWs As Worksheet, ByVal dstAddr As String)
    Dim srcRng As Range
    Set srcRng = srcWs.Range(srcAddr)
    srcRng.Copy
    dstWs.Range(dstAddr).PasteSpecial xlPasteValues
    dstWs.Range(dstAddr).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
End Sub

Public Sub BuildSeverityChart(ws As Worksheet, ByVal anchorAddr As String)
    Dim wsKpi As Worksheet: Set wsKpi = Sheets(SH_KPI)
    Dim anchor As Range: Set anchor = ws.Range(anchorAddr)

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=anchor.Left, Top:=anchor.Top, Width:=380, Height:=240)
    With co.Chart
        .ChartType = xlBarClustered
        .SetSourceData wsKpi.Range("D1:E10")
        .HasTitle = True
        .ChartTitle.Text = "Tickets by Severity"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        On Error Resume Next
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = CLR_ACCENT
        On Error GoTo 0
        .HasLegend = False
    End With
End Sub

Public Sub BuildLocationChart(ws As Worksheet, ByVal anchorAddr As String)
    Dim wsKpi As Worksheet: Set wsKpi = Sheets(SH_KPI)
    Dim anchor As Range: Set anchor = ws.Range(anchorAddr)

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(Left:=anchor.Left, Top:=anchor.Top, Width:=380, Height:=240)
    With co.Chart
        .ChartType = xlBarClustered
        .SetSourceData wsKpi.Range("M1:N11")
        .HasTitle = True
        .ChartTitle.Text = "Top Reporting Locations"
        .ChartTitle.Font.Size = 12
        .ChartTitle.Font.Bold = True
        On Error Resume Next
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = CLR_GOOD
        On Error GoTo 0
        .HasLegend = False
    End With
End Sub

'==============================================================================
' REFRESH BUTTON
'==============================================================================
Public Sub AddRefreshButton(ws As Worksheet)
    On Error Resume Next
    ws.Shapes("btn_refresh").Delete
    On Error GoTo 0

    Dim btn As Shape
    Set btn = ws.Shapes.AddShape(msoShapeRectangle, _
        ws.Range("B1").Left + 8, ws.Range("A1").Top + 4, 130, 28)
    btn.Name = "btn_refresh"
    btn.Fill.ForeColor.RGB = CLR_ACCENT
    btn.Line.Visible = msoFalse
    btn.TextFrame.Characters.Text = "REFRESH"
    With btn.TextFrame.Characters.Font
        .Color = CLR_TEXT
        .Bold = True
        .Size = 11
    End With
    btn.TextFrame.HorizontalAlignment = xlHAlignCenter
    btn.TextFrame.VerticalAlignment = xlVAlignCenter
    btn.OnAction = "BuildDashboard"
End Sub
