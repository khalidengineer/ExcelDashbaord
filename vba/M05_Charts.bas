Attribute VB_Name = "M05_Charts"
'==============================================================================
' MODULE      : M05_Charts
' DESCRIPTION : Chart Rendering Engine.  All charts read from named blocks on
'               the Data_Model sheet (located by section header).  Every chart
'               is restyled to dark NOC aesthetic.
'==============================================================================
Option Explicit
Option Compare Text

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
    M04_DashboardUI.BuildDashboardMain
    M06_Modules.RefreshAllModules
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
    slaVal = M03_KpiEngine.GetKpi("SLA_Compliance")
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
    hidden.Range("AD2").Value = "SLA":             hidden.Range("AE2").Value = NzNum(M03_KpiEngine.GetKpi("SLA_Compliance")) * 100
    hidden.Range("AD3").Value = "Stability":       hidden.Range("AE3").Value = NzNum(M03_KpiEngine.GetKpi("Stability_Score"))
    hidden.Range("AD4").Value = "Closure":         hidden.Range("AE4").Value = NzNum(M03_KpiEngine.GetKpi("ClosureRate")) * 100
    hidden.Range("AD5").Value = "Efficiency":      hidden.Range("AE5").Value = NzNum(M03_KpiEngine.GetKpi("Agent_Efficiency"))
    hidden.Range("AD6").Value = "CSAT":            hidden.Range("AE6").Value = NzNum(M03_KpiEngine.GetKpi("Customer_Satisfaction"))
    hidden.Range("AD7").Value = "Health":          hidden.Range("AE7").Value = NzNum(M03_KpiEngine.GetKpi("OpsHealth_Score"))

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

Public Function NzNum(ByVal v As Variant) As Double
    If IsNumeric(v) Then NzNum = CDbl(v) Else NzNum = 0
End Function

Private Function RandSuffix() As String
    RandSuffix = Format(Now, "hhnnss") & Int(Rnd * 1000)
End Function
