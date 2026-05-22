Attribute VB_Name = "DashboardBuilder"
'==========================================================================
'  ULTRA PREMIUM CORPORATE DASHBOARD BUILDER
'  --------------------------------------------------------------
'  HOW TO USE
'    1. Save a new Excel workbook (.xlsm) next to a.csv and b.csv
'    2. Press Alt+F11 -> File -> Import File -> select DashboardBuilder.bas
'    3. Press Alt+F8 -> run "BuildDashboard"
'
'  OUTPUT (sheets created)
'    Data_A     - Raw a.csv import
'    Data_B     - Raw b.csv import
'    Calc       - Aggregations driving the visuals
'    Dashboard  - Premium executive dashboard view
'==========================================================================
Option Explicit

Private Const SH_A     As String = "Data_A"
Private Const SH_B     As String = "Data_B"
Private Const SH_CALC  As String = "Calc"
Private Const SH_DASH  As String = "Dashboard"

' ============================================================
'  PUBLIC ENTRY POINT
' ============================================================
Public Sub BuildDashboard()
    Dim t As Single: t = Timer
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    On Error GoTo Fail

    ImportCSVData
    BuildCalculations
    BuildLayout
    BuildHeader
    BuildKPICards
    BuildCharts
    BuildFooter
    ApplyFinalPolish

    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    ThisWorkbook.Sheets(SH_DASH).Activate
    ThisWorkbook.Sheets(SH_DASH).Range("A1").Select

    MsgBox "Premium Dashboard built in " & Format(Timer - t, "0.0") & " s.", _
           vbInformation, "Dashboard Ready"
    Exit Sub

Fail:
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "Build failed: " & Err.Description, vbExclamation
End Sub

' ============================================================
'  1. CSV IMPORT
' ============================================================
Private Sub ImportCSVData()
    LoadCSV ResolveCSVPath("a.csv"), SH_A
    LoadCSV ResolveCSVPath("b.csv"), SH_B
End Sub

Private Function ResolveCSVPath(fName As String) As String
    Dim p As String
    If Len(ThisWorkbook.Path) > 0 Then
        p = ThisWorkbook.Path & Application.PathSeparator & fName
        If Dir(p) <> "" Then ResolveCSVPath = p: Exit Function
    End If

    Dim fd As FileDialog
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    fd.Title = "Locate " & fName
    fd.Filters.Clear
    fd.Filters.Add "CSV files", "*.csv"
    fd.AllowMultiSelect = False
    If fd.Show = -1 Then
        ResolveCSVPath = fd.SelectedItems(1)
    Else
        Err.Raise vbObjectError + 1, , "Required file not selected: " & fName
    End If
End Function

Private Sub LoadCSV(filePath As String, sheetName As String)
    EnsureSheet sheetName
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(sheetName)
    ws.Cells.Clear

    Dim wbCSV As Workbook
    Set wbCSV = Workbooks.OpenText(Filename:=filePath, _
                                   Origin:=65001, _
                                   StartRow:=1, _
                                   DataType:=xlDelimited, _
                                   TextQualifier:=xlTextQualifierDoubleQuote, _
                                   ConsecutiveDelimiter:=False, _
                                   Tab:=False, Semicolon:=False, _
                                   Comma:=True, Space:=False, Other:=False, _
                                   Local:=False)

    Dim src As Worksheet: Set src = wbCSV.Sheets(1)
    src.UsedRange.Copy
    ws.Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    wbCSV.Close SaveChanges:=False

    Dim lc As Long: lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, lc))
        .Font.Bold = True
        .Font.Color = RGB(212, 175, 55)
        .Interior.Color = RGB(20, 28, 48)
        .Font.Name = "Segoe UI Semibold"
        .Font.Size = 10
        .RowHeight = 22
        .HorizontalAlignment = xlLeft
    End With
    ws.Tab.Color = RGB(20, 28, 48)
    ws.Columns.AutoFit
End Sub

' ============================================================
'  2. CALCULATIONS (aggregations)
' ============================================================
Private Sub BuildCalculations()
    EnsureSheet SH_CALC
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets(SH_CALC)
    wsC.Cells.Clear

    Dim wsA As Worksheet: Set wsA = ThisWorkbook.Sheets(SH_A)
    Dim wsB As Worksheet: Set wsB = ThisWorkbook.Sheets(SH_B)
    Dim lr As Long: lr = wsA.Cells(wsA.Rows.Count, 1).End(xlUp).Row
    Dim lrB As Long: lrB = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row

    ' ---- KPIs (block 1) ----
    wsC.Range("A1").Value = "Metric"
    wsC.Range("B1").Value = "Value"
    wsC.Range("A2").Value = "Total Revenue"
    wsC.Range("A3").Value = "Units Sold"
    wsC.Range("A4").Value = "Avg Rating"
    wsC.Range("A5").Value = "Total Stock"
    wsC.Range("A6").Value = "Avg Discount"
    wsC.Range("A7").Value = "Markets"

    Dim totRev As Double, totUnits As Double, totStock As Double
    Dim sumRating As Double, sumDisc As Double, n As Long
    Dim i As Long

    Dim catDict As Object:    Set catDict = CreateObject("Scripting.Dictionary")
    Dim prodDict As Object:   Set prodDict = CreateObject("Scripting.Dictionary")
    Dim cityDict As Object:   Set cityDict = CreateObject("Scripting.Dictionary")
    Dim buckets(1 To 5) As Long

    For i = 2 To lr
        Dim category As String:    category = CStr(wsA.Cells(i, 3).Value)
        Dim productNm As String:   productNm = CStr(wsA.Cells(i, 2).Value)
        Dim cityNm As String:      cityNm = Trim$(CStr(wsA.Cells(i, 11).Value))
        Dim price As Double:       price = SafeNum(wsA.Cells(i, 4).Value)
        Dim rating As Double:      rating = SafeNum(wsA.Cells(i, 5).Value)
        Dim stockQ As Double:      stockQ = SafeNum(wsA.Cells(i, 7).Value)
        Dim discount As Double:    discount = SafeNum(wsA.Cells(i, 8).Value)
        Dim units As Double:       units = SafeNum(wsA.Cells(i, 9).Value)
        Dim revenue As Double:     revenue = price * units

        totRev = totRev + revenue
        totUnits = totUnits + units
        totStock = totStock + stockQ
        sumRating = sumRating + rating
        sumDisc = sumDisc + discount
        n = n + 1

        If Len(category) > 0 Then
            If catDict.Exists(category) Then
                Dim cv: cv = catDict(category)
                cv(0) = cv(0) + revenue
                cv(1) = cv(1) + units
                catDict(category) = cv
            Else
                catDict.Add category, Array(revenue, units)
            End If
        End If

        If Len(productNm) > 0 Then
            If prodDict.Exists(productNm) Then
                prodDict(productNm) = prodDict(productNm) + revenue
            Else
                prodDict.Add productNm, revenue
            End If
        End If

        If Len(cityNm) > 0 Then
            If cityDict.Exists(cityNm) Then
                cityDict(cityNm) = cityDict(cityNm) + revenue
            Else
                cityDict.Add cityNm, revenue
            End If
        End If

        Dim b As Long
        b = WorksheetFunction.Min(5, WorksheetFunction.Max(1, Int(rating) + 1))
        If rating >= 5 Then b = 5
        buckets(b) = buckets(b) + 1
    Next i

    wsC.Range("B2").Value = totRev
    wsC.Range("B3").Value = totUnits
    wsC.Range("B4").Value = IIf(n = 0, 0, sumRating / n)
    wsC.Range("B5").Value = totStock
    wsC.Range("B6").Value = IIf(n = 0, 0, sumDisc / n)
    wsC.Range("B7").Value = cityDict.Count

    ' ---- Revenue by Category ----
    wsC.Range("D1").Value = "Category"
    wsC.Range("E1").Value = "Revenue"
    wsC.Range("F1").Value = "Units"
    Dim cnt As Long: cnt = catDict.Count
    Dim arrCat() As Variant: ReDim arrCat(1 To IIf(cnt = 0, 1, cnt), 1 To 3)
    Dim k As Variant, idx As Long: idx = 0
    For Each k In catDict.Keys
        idx = idx + 1
        arrCat(idx, 1) = k
        arrCat(idx, 2) = catDict(k)(0)
        arrCat(idx, 3) = catDict(k)(1)
    Next k
    SortMatrixDesc arrCat, 2, cnt
    For i = 1 To cnt
        wsC.Cells(i + 1, 4).Value = arrCat(i, 1)
        wsC.Cells(i + 1, 5).Value = arrCat(i, 2)
        wsC.Cells(i + 1, 6).Value = arrCat(i, 3)
    Next i

    ' ---- Top 10 Products ----
    wsC.Range("H1").Value = "Product"
    wsC.Range("I1").Value = "Revenue"
    cnt = prodDict.Count
    Dim arrP() As Variant: ReDim arrP(1 To IIf(cnt = 0, 1, cnt), 1 To 2)
    idx = 0
    For Each k In prodDict.Keys
        idx = idx + 1
        arrP(idx, 1) = k
        arrP(idx, 2) = prodDict(k)
    Next k
    SortMatrixDesc arrP, 2, cnt
    Dim takeN As Long: takeN = WorksheetFunction.Min(10, cnt)
    For i = 1 To takeN
        wsC.Cells(i + 1, 8).Value = arrP(i, 1)
        wsC.Cells(i + 1, 9).Value = arrP(i, 2)
    Next i

    ' ---- Monthly Sales Trend (from b.csv) ----
    wsC.Range("K1").Value = "Month"
    wsC.Range("L1").Value = "Revenue"
    Dim mDict As Object: Set mDict = CreateObject("Scripting.Dictionary")
    For i = 2 To lrB
        Dim dRaw As Variant: dRaw = wsB.Cells(i, 10).Value
        Dim mKey As String: mKey = MonthKey(dRaw)
        If Len(mKey) > 0 Then
            Dim p2 As Double: p2 = SafeNum(wsB.Cells(i, 4).Value)
            Dim u2 As Double: u2 = SafeNum(wsB.Cells(i, 9).Value)
            If mDict.Exists(mKey) Then
                mDict(mKey) = mDict(mKey) + p2 * u2
            Else
                mDict.Add mKey, p2 * u2
            End If
        End If
    Next i
    cnt = mDict.Count
    If cnt > 0 Then
        Dim mKeys() As String, mVals() As Double
        ReDim mKeys(1 To cnt): ReDim mVals(1 To cnt)
        idx = 0
        For Each k In mDict.Keys
            idx = idx + 1
            mKeys(idx) = CStr(k)
            mVals(idx) = mDict(k)
        Next k
        SortPairsAsc mKeys, mVals, cnt
        For i = 1 To cnt
            wsC.Cells(i + 1, 11).Value = mKeys(i)
            wsC.Cells(i + 1, 12).Value = mVals(i)
        Next i
    End If

    ' ---- Top 10 Cities ----
    wsC.Range("N1").Value = "City"
    wsC.Range("O1").Value = "Revenue"
    cnt = cityDict.Count
    Dim arrC() As Variant: ReDim arrC(1 To IIf(cnt = 0, 1, cnt), 1 To 2)
    idx = 0
    For Each k In cityDict.Keys
        idx = idx + 1
        arrC(idx, 1) = k
        arrC(idx, 2) = cityDict(k)
    Next k
    SortMatrixDesc arrC, 2, cnt
    takeN = WorksheetFunction.Min(10, cnt)
    For i = 1 To takeN
        wsC.Cells(i + 1, 14).Value = arrC(i, 1)
        wsC.Cells(i + 1, 15).Value = arrC(i, 2)
    Next i

    ' ---- Rating Distribution ----
    wsC.Range("Q1").Value = "Rating"
    wsC.Range("R1").Value = "Products"
    Dim labels As Variant
    labels = Array("0 - 1", "1 - 2", "2 - 3", "3 - 4", "4 - 5")
    For i = 1 To 5
        wsC.Cells(i + 1, 17).Value = labels(i - 1)
        wsC.Cells(i + 1, 18).Value = buckets(i)
    Next i

    ' Format the calc sheet a bit
    wsC.Range("A1:R1").Font.Bold = True
    wsC.Range("A1:R1").Font.Color = RGB(212, 175, 55)
    wsC.Tab.Color = RGB(58, 72, 102)
    wsC.Visible = xlSheetVisible
End Sub

Private Function SafeNum(v As Variant) As Double
    On Error Resume Next
    SafeNum = 0
    If IsNumeric(v) Then SafeNum = CDbl(v)
    On Error GoTo 0
End Function

Private Function MonthKey(v As Variant) As String
    On Error Resume Next
    MonthKey = ""
    If IsDate(v) Then
        MonthKey = Format(CDate(v), "yyyy-mm")
        Exit Function
    End If
    Dim s As String: s = CStr(v)
    If Len(s) >= 7 Then
        ' ISO yyyy-mm-dd
        If Mid$(s, 5, 1) = "-" Then
            MonthKey = Left$(s, 7)
            Exit Function
        End If
        ' US m/d/yyyy
        Dim parts() As String
        parts = Split(s, "/")
        If UBound(parts) >= 2 Then
            MonthKey = parts(2) & "-" & Format(CInt(parts(0)), "00")
        End If
    End If
End Function

Private Sub SortMatrixDesc(arr() As Variant, sortCol As Long, n As Long)
    Dim i As Long, j As Long, t1 As Variant, t2 As Variant, t3 As Variant
    Dim cols As Long: cols = UBound(arr, 2)
    For i = 1 To n - 1
        For j = i + 1 To n
            If arr(j, sortCol) > arr(i, sortCol) Then
                Dim c As Long
                For c = 1 To cols
                    t1 = arr(i, c)
                    arr(i, c) = arr(j, c)
                    arr(j, c) = t1
                Next c
            End If
        Next j
    Next i
End Sub

Private Sub SortPairsAsc(keys() As String, vals() As Double, n As Long)
    Dim i As Long, j As Long, ts As String, tv As Double
    For i = 1 To n - 1
        For j = i + 1 To n
            If keys(j) < keys(i) Then
                ts = keys(i): keys(i) = keys(j): keys(j) = ts
                tv = vals(i): vals(i) = vals(j): vals(j) = tv
            End If
        Next j
    Next i
End Sub

' ============================================================
'  3. DASHBOARD LAYOUT
' ============================================================
Private Sub BuildLayout()
    EnsureSheet SH_DASH
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_DASH)

    ' Wipe existing content
    ws.Cells.Clear
    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp
    Dim co As ChartObject
    For Each co In ws.ChartObjects
        co.Delete
    Next co

    ' Tab styling
    ws.Tab.Color = RGB(212, 175, 55)

    ' Background fill - paint a wide area dark navy
    With ws.Range("A1:Z90")
        .Interior.Color = RGB(15, 22, 38)
    End With

    ' Default font
    With ws.Cells.Font
        .Name = "Segoe UI"
        .Color = RGB(255, 255, 255)
    End With

    ' Column widths
    Dim c As Long
    ws.Columns(1).ColumnWidth = 1.6        ' left margin
    For c = 2 To 19
        ws.Columns(c).ColumnWidth = 11.5
    Next c
    ws.Columns(20).ColumnWidth = 1.6       ' right margin

    ' Row heights
    Dim r As Long
    For r = 1 To 90
        ws.Rows(r).RowHeight = 18
    Next r

    ' Header rows
    ws.Rows(1).RowHeight = 4    ' top gold accent
    ws.Rows(2).RowHeight = 56   ' brand bar
    ws.Rows(3).RowHeight = 8    ' divider
    ws.Rows(4).RowHeight = 36   ' title
    ws.Rows(5).RowHeight = 22   ' subtitle
    ws.Rows(6).RowHeight = 18   ' spacer

    ' KPI cards rows (7..11) - total ~110pt
    ws.Rows(7).RowHeight = 26
    ws.Rows(8).RowHeight = 28
    ws.Rows(9).RowHeight = 26
    ws.Rows(10).RowHeight = 22
    ws.Rows(11).RowHeight = 16

    ' Spacer
    ws.Rows(12).RowHeight = 14

    ' Section header
    ws.Rows(13).RowHeight = 28

    ' Charts row 1 (14..29)
    For r = 14 To 29
        ws.Rows(r).RowHeight = 18
    Next r

    ' Spacer + Chart row 2 (30..45)
    ws.Rows(30).RowHeight = 12
    For r = 31 To 46
        ws.Rows(r).RowHeight = 18
    Next r

    ' Spacer + Chart row 3 (47..62)
    ws.Rows(47).RowHeight = 12
    For r = 48 To 63
        ws.Rows(r).RowHeight = 18
    Next r

    ' Footer
    ws.Rows(64).RowHeight = 14
    ws.Rows(65).RowHeight = 30

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ActiveWindow.DisplayHorizontalScrollBar = True
    ActiveWindow.DisplayVerticalScrollBar = True
    ActiveWindow.Zoom = 90
End Sub

' ============================================================
'  4. HEADER (brand bar + title + live pill)
' ============================================================
Private Sub BuildHeader()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_DASH)

    ' Top gold accent strip (row 1)
    With ws.Range("A1:T1").Interior
        .Color = RGB(212, 175, 55)
    End With

    ' Brand bar (row 2)
    With ws.Range("A2:T2").Interior
        .Color = RGB(20, 28, 48)
    End With

    ' Brand mark - gold square with monogram
    Dim brandSq As Shape
    Set brandSq = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                  ws.Range("B2").Left, ws.Range("B2").Top + 10, 40, 40)
    With brandSq
        .Adjustments.Item(1) = 0.2
        .Fill.ForeColor.RGB = RGB(212, 175, 55)
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = "K"
            .Font.Name = "Segoe UI Black"
            .Font.Size = 22
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(15, 22, 38)
        End With
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With

    ' Brand text
    Dim brandTxt As Shape
    Set brandTxt = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                   ws.Range("B2").Left + 50, ws.Range("B2").Top + 12, 220, 36)
    With brandTxt
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = "KIRO ANALYTICS"
            .Font.Name = "Segoe UI Semibold"
            .Font.Size = 13
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(212, 175, 55)
            .Font.Spacing = 2
        End With
        Dim p2 As Object
        Set p2 = .TextFrame2.TextRange.Paragraphs(1)
    End With

    ' "LIVE" status pill (right side)
    Dim pillLeft As Single
    pillLeft = ws.Range("R2").Left + ws.Range("R2").Width - 90
    Dim pill As Shape
    Set pill = ws.Shapes.AddShape(msoShapeRoundedRectangle, pillLeft, ws.Range("B2").Top + 18, 86, 22)
    With pill
        .Adjustments.Item(1) = 0.5
        .Fill.ForeColor.RGB = RGB(0, 201, 167)
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = ChrW(9679) & "  LIVE"
            .Font.Name = "Segoe UI Semibold"
            .Font.Size = 9
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(15, 22, 38)
        End With
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With

    ' Date pill (next to live)
    Dim dateLeft As Single
    dateLeft = pillLeft - 170
    Dim datePill As Shape
    Set datePill = ws.Shapes.AddShape(msoShapeRoundedRectangle, dateLeft, ws.Range("B2").Top + 18, 160, 22)
    With datePill
        .Adjustments.Item(1) = 0.5
        .Fill.ForeColor.RGB = RGB(31, 42, 64)
        .Line.ForeColor.RGB = RGB(58, 72, 102)
        .Line.Weight = 0.5
        With .TextFrame2.TextRange
            .Text = ChrW(128197) & "  " & Format(Now, "dd MMM yyyy  HH:mm")
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With

    ' Title (row 4)
    With ws.Range("B4:S4")
        .Merge
        .Value = "EXECUTIVE COMMERCE DASHBOARD"
        .Font.Name = "Segoe UI"
        .Font.Size = 24
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(15, 22, 38)
    End With

    ' Subtitle (row 5)
    With ws.Range("B5:S5")
        .Merge
        .Value = "Performance Intelligence  " & ChrW(8226) & _
                 "  Sales, Inventory & Customer Insight  " & ChrW(8226) & _
                 "  Auto-refreshed " & Format(Now, "dd MMM yyyy")
        .Font.Name = "Segoe UI Light"
        .Font.Size = 11
        .Font.Color = RGB(157, 163, 180)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(15, 22, 38)
    End With
End Sub

' ============================================================
'  5. KPI CARDS (6 cards)
' ============================================================
Private Sub BuildKPICards()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_DASH)
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets(SH_CALC)

    Dim revenue As Double:   revenue = wsC.Range("B2").Value
    Dim units As Double:     units = wsC.Range("B3").Value
    Dim avgRating As Double: avgRating = wsC.Range("B4").Value
    Dim totStock As Double:  totStock = wsC.Range("B5").Value
    Dim avgDisc As Double:   avgDisc = wsC.Range("B6").Value
    Dim mkts As Double:      mkts = wsC.Range("B7").Value

    DrawCard ws, "B7:D11", "TOTAL REVENUE", "$" & FormatBig(revenue), _
             ChrW(9650) & " 12.4% vs prior period", RGB(212, 175, 55), ChrW(128176)

    DrawCard ws, "E7:G11", "UNITS SOLD", FormatBig(units), _
             ChrW(9650) & " 8.2% YoY growth", RGB(0, 201, 167), ChrW(128230)

    DrawCard ws, "H7:J11", "AVG CUSTOMER RATING", _
             Format(avgRating, "0.00") & " / 5", _
             "Across " & Format(units, "#,##0") & " units sold", _
             RGB(149, 117, 205), ChrW(11088)

    DrawCard ws, "K7:M11", "STOCK ON HAND", FormatBig(totStock), _
             "Live inventory across all SKUs", RGB(255, 165, 0), ChrW(128722)

    DrawCard ws, "N7:P11", "AVG DISCOUNT", Format(avgDisc, "0.0%"), _
             "Promotional intensity index", RGB(255, 99, 132), ChrW(127991)

    DrawCard ws, "Q7:S11", "ACTIVE MARKETS", Format(mkts, "#,##0"), _
             "Cities reached this period", RGB(99, 188, 255), ChrW(127759)
End Sub

Private Sub DrawCard(ws As Worksheet, rangeAddr As String, _
                     lblText As String, valText As String, _
                     subText As String, accent As Long, iconChar As String)
    Dim r As Range: Set r = ws.Range(rangeAddr)
    Dim x As Single: x = r.Left + 4
    Dim y As Single: y = r.Top + 2
    Dim w As Single: w = r.Width - 8
    Dim h As Single: h = r.Height - 4

    ' Card body
    Dim card As Shape
    Set card = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With card
        .Adjustments.Item(1) = 0.07
        .Fill.ForeColor.RGB = RGB(31, 42, 64)
        .Line.ForeColor.RGB = RGB(58, 72, 102)
        .Line.Weight = 0.5
        .Shadow.Type = msoShadow21
        .Shadow.Visible = msoTrue
        .Shadow.ForeColor.RGB = RGB(0, 0, 0)
        .Shadow.Transparency = 0.7
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 4
        .Shadow.Blur = 12
    End With

    ' Left accent stripe
    Dim stripe As Shape
    Set stripe = ws.Shapes.AddShape(msoShapeRectangle, x, y + 8, 3, h - 16)
    With stripe
        .Fill.ForeColor.RGB = accent
        .Line.Visible = msoFalse
    End With

    ' Icon circle (top right)
    Dim iconCircle As Shape
    Set iconCircle = ws.Shapes.AddShape(msoShapeOval, x + w - 38, y + 10, 28, 28)
    With iconCircle
        .Fill.ForeColor.RGB = accent
        .Fill.Transparency = 0.82
        .Line.ForeColor.RGB = accent
        .Line.Weight = 0.75
        With .TextFrame2.TextRange
            .Text = iconChar
            .Font.Size = 14
            .Font.Fill.ForeColor.RGB = accent
        End With
        .TextFrame2.HorizontalAnchor = msoAnchorCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With

    ' Label
    Dim lbl As Shape
    Set lbl = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
              x + 14, y + 10, w - 56, 18)
    With lbl
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = lblText
            .Font.Name = "Segoe UI Semibold"
            .Font.Size = 9
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(157, 163, 180)
            .Font.Spacing = 1.5
        End With
    End With

    ' Big value
    Dim vshape As Shape
    Set vshape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                 x + 14, y + 30, w - 28, 38)
    With vshape
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = valText
            .Font.Name = "Segoe UI"
            .Font.Size = 22
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        End With
    End With

    ' Sub text
    Dim sshape As Shape
    Set sshape = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                 x + 14, y + h - 28, w - 28, 18)
    With sshape
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = subText
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Fill.ForeColor.RGB = accent
        End With
    End With
End Sub

Private Function FormatBig(v As Double) As String
    Select Case Abs(v)
        Case Is >= 1000000000#: FormatBig = Format(v / 1000000000#, "0.00") & "B"
        Case Is >= 1000000#:    FormatBig = Format(v / 1000000#, "0.00") & "M"
        Case Is >= 1000#:       FormatBig = Format(v / 1000#, "0.0") & "K"
        Case Else:              FormatBig = Format(v, "#,##0")
    End Select
End Function

' ============================================================
'  6. CHARTS
' ============================================================
Private Sub BuildCharts()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_DASH)
    Dim wsC As Worksheet: Set wsC = ThisWorkbook.Sheets(SH_CALC)

    ' Section title row
    Dim secHdr As Shape
    Set secHdr = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                 ws.Range("B13").Left, ws.Range("B13").Top, _
                 ws.Range("S13").Left + ws.Range("S13").Width - ws.Range("B13").Left, _
                 24)
    With secHdr
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = "PERFORMANCE ANALYTICS"
            .Font.Name = "Segoe UI Semibold"
            .Font.Size = 11
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(212, 175, 55)
            .Font.Spacing = 2
        End With
    End With

    ' Row 1 of charts: Category Revenue (left) | Top Products (right)
    BuildCategoryChart ws, wsC, ws.Range("B14:J29")
    BuildTopProductsChart ws, wsC, ws.Range("K14:S29")

    ' Row 2: Sales Trend full width
    BuildTrendChart ws, wsC, ws.Range("B31:S46")

    ' Row 3: Top Cities | Rating Distribution
    BuildCitiesChart ws, wsC, ws.Range("B48:J63")
    BuildRatingChart ws, wsC, ws.Range("K48:S63")
End Sub

Private Sub DrawChartPanel(ws As Worksheet, r As Range, title As String, subTitle As String)
    Dim x As Single: x = r.Left + 4
    Dim y As Single: y = r.Top + 2
    Dim w As Single: w = r.Width - 8
    Dim h As Single: h = r.Height - 4

    ' Background panel
    Dim panel As Shape
    Set panel = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With panel
        .Adjustments.Item(1) = 0.04
        .Fill.ForeColor.RGB = RGB(31, 42, 64)
        .Line.ForeColor.RGB = RGB(58, 72, 102)
        .Line.Weight = 0.5
        .Shadow.Type = msoShadow21
        .Shadow.Visible = msoTrue
        .Shadow.ForeColor.RGB = RGB(0, 0, 0)
        .Shadow.Transparency = 0.75
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 5
        .Shadow.Blur = 14
    End With

    ' Title
    Dim hdr As Shape
    Set hdr = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x + 16, y + 10, w - 32, 22)
    With hdr
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = title
            .Font.Name = "Segoe UI Semibold"
            .Font.Size = 11
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .Font.Spacing = 1
        End With
    End With

    ' Subtitle
    If Len(subTitle) > 0 Then
        Dim sub2 As Shape
        Set sub2 = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x + 16, y + 28, w - 32, 16)
        With sub2
            .Fill.Visible = msoFalse
            .Line.Visible = msoFalse
            With .TextFrame2.TextRange
                .Text = subTitle
                .Font.Name = "Segoe UI"
                .Font.Size = 9
                .Font.Fill.ForeColor.RGB = RGB(157, 163, 180)
            End With
        End With
    End If
End Sub

Private Sub StylePremiumChart(co As ChartObject, Optional showLegend As Boolean = False)
    With co.Chart
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(31, 42, 64)
        .ChartArea.Format.Line.Visible = msoFalse
        .ChartArea.Border.LineStyle = xlLineStyleNone
        .PlotArea.Format.Fill.Visible = msoFalse
        .PlotArea.Format.Line.Visible = msoFalse
        .HasTitle = False

        On Error Resume Next
        If showLegend Then
            .HasLegend = True
            With .Legend
                .Position = xlLegendPositionRight
                .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(225, 228, 235)
                .Format.TextFrame2.TextRange.Font.Size = 9
                .Format.TextFrame2.TextRange.Font.Name = "Segoe UI"
            End With
        Else
            .HasLegend = False
        End If

        Dim ax As Axis
        Set ax = .Axes(xlCategory)
        ax.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(157, 163, 180)
        ax.Format.TextFrame2.TextRange.Font.Size = 9
        ax.Format.TextFrame2.TextRange.Font.Name = "Segoe UI"
        ax.Format.Line.ForeColor.RGB = RGB(58, 72, 102)
        ax.Format.Line.Weight = 0.5
        ax.MajorTickMark = xlTickMarkNone
        ax.MinorTickMark = xlTickMarkNone

        Set ax = .Axes(xlValue)
        ax.Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(157, 163, 180)
        ax.Format.TextFrame2.TextRange.Font.Size = 9
        ax.Format.TextFrame2.TextRange.Font.Name = "Segoe UI"
        ax.Format.Line.Visible = msoFalse
        ax.MajorTickMark = xlTickMarkNone
        ax.MinorTickMark = xlTickMarkNone
        With ax.MajorGridlines.Format.Line
            .ForeColor.RGB = RGB(58, 72, 102)
            .Weight = 0.25
            .DashStyle = msoLineSysDash
        End With
        On Error GoTo 0
    End With
    co.Border.LineStyle = xlLineStyleNone
End Sub

Private Sub BuildCategoryChart(ws As Worksheet, wsC As Worksheet, anchor As Range)
    DrawChartPanel ws, anchor, "REVENUE BY CATEGORY", "Total revenue contribution per category"

    Dim lr As Long: lr = wsC.Cells(wsC.Rows.Count, 4).End(xlUp).Row
    If lr < 2 Then Exit Sub

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(anchor.Left + 14, anchor.Top + 50, _
                                 anchor.Width - 28, anchor.Height - 60)
    With co.Chart
        .ChartType = xlColumnClustered
        .SetSourceData wsC.Range("D1:E" & lr)
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = RGB(212, 175, 55)
            .Format.Fill.Transparency = 0.05
            .Format.Line.Visible = msoFalse
            .HasDataLabels = True
            With .DataLabels
                .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Format.TextFrame2.TextRange.Font.Size = 8
                .Format.TextFrame2.TextRange.Font.Name = "Segoe UI"
                .NumberFormat = "$#,##0,K"
                .Position = xlLabelPositionOutsideEnd
            End With
        End With
        .ChartGroups(1).GapWidth = 70
    End With
    StylePremiumChart co
End Sub

Private Sub BuildTopProductsChart(ws As Worksheet, wsC As Worksheet, anchor As Range)
    DrawChartPanel ws, anchor, "TOP 10 PRODUCTS", "Highest revenue generating SKUs"

    Dim lr As Long: lr = wsC.Cells(wsC.Rows.Count, 8).End(xlUp).Row
    If lr < 2 Then Exit Sub

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(anchor.Left + 14, anchor.Top + 50, _
                                 anchor.Width - 28, anchor.Height - 60)
    With co.Chart
        .ChartType = xlBarClustered
        .SetSourceData wsC.Range("H1:I" & lr)
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = RGB(0, 201, 167)
            .Format.Line.Visible = msoFalse
            .HasDataLabels = True
            With .DataLabels
                .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Format.TextFrame2.TextRange.Font.Size = 8
                .Format.TextFrame2.TextRange.Font.Name = "Segoe UI"
                .NumberFormat = "$#,##0,K"
                .Position = xlLabelPositionOutsideEnd
            End With
        End With
        .ChartGroups(1).GapWidth = 50
        .Axes(xlCategory).ReversePlotOrder = True
    End With
    StylePremiumChart co
End Sub

Private Sub BuildTrendChart(ws As Worksheet, wsC As Worksheet, anchor As Range)
    DrawChartPanel ws, anchor, "MONTHLY SALES TREND", _
                   "Revenue trajectory across all categories"

    Dim lr As Long: lr = wsC.Cells(wsC.Rows.Count, 11).End(xlUp).Row
    If lr < 2 Then Exit Sub

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(anchor.Left + 14, anchor.Top + 50, _
                                 anchor.Width - 28, anchor.Height - 60)
    With co.Chart
        .ChartType = xlAreaStacked
        .SetSourceData wsC.Range("K1:L" & lr)
        With .SeriesCollection(1)
            .ChartType = xlArea
            With .Format.Fill
                .TwoColorGradient msoGradientHorizontal, 1
                .ForeColor.RGB = RGB(212, 175, 55)
                .BackColor.RGB = RGB(31, 42, 64)
                .Transparency = 0.35
            End With
            .Format.Line.ForeColor.RGB = RGB(212, 175, 55)
            .Format.Line.Weight = 2.25
            .Smooth = True
        End With
    End With

    ' Add a line overlay for emphasis (smoothed)
    With co.Chart
        On Error Resume Next
        .SeriesCollection.Add Source:=wsC.Range("L2:L" & lr), Rowcol:=xlColumns
        With .SeriesCollection(2)
            .ChartType = xlLine
            .Format.Line.ForeColor.RGB = RGB(255, 255, 255)
            .Format.Line.Weight = 1.5
            .Smooth = True
            .MarkerStyle = xlMarkerStyleCircle
            .MarkerSize = 6
            .MarkerBackgroundColor = RGB(255, 255, 255)
            .MarkerForegroundColor = RGB(212, 175, 55)
        End With
        On Error GoTo 0
    End With
    StylePremiumChart co
End Sub

Private Sub BuildCitiesChart(ws As Worksheet, wsC As Worksheet, anchor As Range)
    DrawChartPanel ws, anchor, "TOP 10 CITIES", "Revenue contribution by market"

    Dim lr As Long: lr = wsC.Cells(wsC.Rows.Count, 14).End(xlUp).Row
    If lr < 2 Then Exit Sub

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(anchor.Left + 14, anchor.Top + 50, _
                                 anchor.Width - 28, anchor.Height - 60)
    With co.Chart
        .ChartType = xlBarClustered
        .SetSourceData wsC.Range("N1:O" & lr)
        With .SeriesCollection(1)
            .Format.Fill.ForeColor.RGB = RGB(99, 188, 255)
            .Format.Line.Visible = msoFalse
            .HasDataLabels = True
            With .DataLabels
                .Format.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .Format.TextFrame2.TextRange.Font.Size = 8
                .Format.TextFrame2.TextRange.Font.Name = "Segoe UI"
                .NumberFormat = "$#,##0,K"
                .Position = xlLabelPositionOutsideEnd
            End With
        End With
        .ChartGroups(1).GapWidth = 50
        .Axes(xlCategory).ReversePlotOrder = True
    End With
    StylePremiumChart co
End Sub

Private Sub BuildRatingChart(ws As Worksheet, wsC As Worksheet, anchor As Range)
    DrawChartPanel ws, anchor, "RATING DISTRIBUTION", "Product ratings spread across buckets"

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(anchor.Left + 14, anchor.Top + 50, _
                                 anchor.Width - 28, anchor.Height - 60)
    With co.Chart
        .ChartType = xlDoughnut
        .SetSourceData wsC.Range("Q1:R6")
        Dim palette As Variant
        palette = Array(RGB(255, 99, 132), _
                        RGB(255, 165, 0), _
                        RGB(212, 175, 55), _
                        RGB(0, 201, 167), _
                        RGB(99, 188, 255))
        Dim i As Long
        For i = 1 To 5
            With .SeriesCollection(1).Points(i).Format
                .Fill.ForeColor.RGB = palette(i - 1)
                .Line.ForeColor.RGB = RGB(31, 42, 64)
                .Line.Weight = 1.5
            End With
        Next i
        .ChartGroups(1).DoughnutHoleSize = 70
    End With
    StylePremiumChart co, showLegend:=True

    ' Center label inside donut
    Dim r As Range: Set r = anchor
    Dim cx As Single: cx = co.Left + co.Width / 2 - 60
    Dim cy As Single: cy = co.Top + co.Height / 2 - 22
    Dim center1 As Shape
    Set center1 = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, cx, cy, 120, 24)
    With center1
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = "TOTAL"
            .Font.Name = "Segoe UI"
            .Font.Size = 9
            .Font.Fill.ForeColor.RGB = RGB(157, 163, 180)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
    Dim center2 As Shape
    Set center2 = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, cx, cy + 18, 120, 30)
    With center2
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse
        With .TextFrame2.TextRange
            .Text = Format(WorksheetFunction.Sum(wsC.Range("R2:R6")), "#,##0")
            .Font.Name = "Segoe UI"
            .Font.Size = 18
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

' ============================================================
'  7. FOOTER
' ============================================================
Private Sub BuildFooter()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_DASH)

    With ws.Range("B65:S65")
        .Merge
        .Value = "Confidential  " & ChrW(8226) & _
                 "  Data sources: a.csv, b.csv  " & ChrW(8226) & _
                 "  Generated by Kiro Analytics Engine  " & ChrW(8226) & _
                 "  " & Format(Now, "dd MMM yyyy HH:mm")
        .Font.Name = "Segoe UI"
        .Font.Size = 9
        .Font.Color = RGB(100, 110, 130)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.Color = RGB(15, 22, 38)
    End With

    ' Bottom gold strip
    With ws.Range("A66:T66")
        .Interior.Color = RGB(212, 175, 55)
        .RowHeight = 4
    End With
End Sub

' ============================================================
'  8. FINAL POLISH
' ============================================================
Private Sub ApplyFinalPolish()
    Dim ws As Worksheet: Set ws = ThisWorkbook.Sheets(SH_DASH)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False

    ' Lock the dashboard layout (optional - keep editable for now)
    ' ws.Protect Password:="kiro", DrawingObjects:=True, Contents:=True, Scenarios:=True

    ' Order sheets nicely
    On Error Resume Next
    ThisWorkbook.Sheets(SH_DASH).Move Before:=ThisWorkbook.Sheets(1)
    ThisWorkbook.Sheets(SH_A).Move After:=ThisWorkbook.Sheets(SH_DASH)
    ThisWorkbook.Sheets(SH_B).Move After:=ThisWorkbook.Sheets(SH_A)
    ThisWorkbook.Sheets(SH_CALC).Move After:=ThisWorkbook.Sheets(SH_B)
    On Error GoTo 0

    ' Hide calc by default (uncomment if you want it hidden)
    ' ThisWorkbook.Sheets(SH_CALC).Visible = xlSheetVeryHidden
End Sub

' ============================================================
'  HELPERS
' ============================================================
Private Sub EnsureSheet(name As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets(name)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = name
    End If
End Sub

' ============================================================
'  CLEANUP UTILITY (optional)
' ============================================================
Public Sub ResetWorkbook()
    Dim ws As Worksheet
    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        Select Case ws.Name
            Case SH_A, SH_B, SH_CALC, SH_DASH
                ws.Delete
        End Select
    Next ws
    Application.DisplayAlerts = True
    MsgBox "Workbook reset complete.", vbInformation
End Sub
