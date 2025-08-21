Attribute VB_Name = "Plot_MobileBand_XYChart"
Const BandStep As Integer = 5, BandFirst As Integer = 1, BandLast As Integer = 80 'Avoid larger than 66',
Const BandNumber1 As Integer = 1, BandNumber2 As Integer = 2
Const UplinkMinFreq As Integer = 3, UplinkMaxFreq As Integer = 4
Const DownlinkMinFreq As Integer = 5, DownlinkMaxFreq As Integer = 6
Const FreqStep As Integer = 500, FreqMin As Integer = 0, FreqMax As Integer = 2700
Const UplinkChart As String = "Chart 1", DownlinkChart As String = "Chart 2"
Const LineThinkness As Integer = 4
Const UplinkBandTitle As String = "LTE & 5GNR Band", DownlinkBandTitle As String = "LTE & 5GNR Band", TitleRatio As Integer = 30
Const DataSource As String = "LTE_NR"   'Define the chart source worksheet name
Const ChartGap As Integer = 5, ChartHeight As Integer = 5 * 150, ChartWidth As Integer = (4 + 1) * 120 * 2
Const ChartTop As Integer = 0, ChartLeft As Integer = 500
'msoLineCapFlat
'.SendKeys "F{RETURN}{ENTER}" ' Flat

Sub Plot_LTE_5GNR_Uplink_Downlink()

Call Plot_LTE_Band_Uplink
Call Plot_LTE_Band_Downlink

End Sub
Sub Plot_LTE_Band_Uplink()

'
' Plot_NR ????X
'
'
Dim Band As Integer
Dim BandHighest As Integer
      
Call CreateChart(UplinkChart)
      
'Update the serie data of the chart source
For Band = BandFirst To BandLast
    ActiveSheet.ChartObjects(UplinkChart).Activate
    ActiveChart.SeriesCollection.NewSeries
    If Worksheets(DataSource).Cells(Band, UplinkMinFreq) <> "N/A" Then
        With ActiveChart.FullSeriesCollection(Band)
            .Name = Worksheets(DataSource).Cells(Band, BandNumber1)
            .XValues = Worksheets(DataSource).Range(Cells(Band, UplinkMinFreq), Cells(Band, UplinkMaxFreq))
            .Values = Worksheets(DataSource).Range(Cells(Band, BandNumber1), Cells(Band, BandNumber2))
        End With
        
        'Update the chart line format and color
        'ActiveChart.FullSeriesCollection(Band).Select
        'With Selection.Format.Line
        With ActiveChart.FullSeriesCollection(Band).Format.Line
            .Visible = msoTrue
            If Cells(Band, 7) = "FDD" And Cells(Band, 8) <> "LTE" And Cells(Band, 9) = "NR" Then
                 .ForeColor.RGB = RGB(0, 0, 255) 'FDD mode, NR only color in Blue
            ElseIf Cells(Band, 7) = "FDD" And Cells(Band, 8) = "LTE" And Cells(Band, 9) <> "NR" Then
                 .ForeColor.RGB = RGB(0, 255, 0) 'FDD mode, LTE only mode color in Green
            ElseIf Cells(Band, 7) = "FDD" And Cells(Band, 8) = "LTE" And Cells(Band, 9) = "NR" Then
                 .ForeColor.RGB = RGB(0, 255, 255) 'FDD mode, LTE and NR mode color in Aqua
            ElseIf Cells(Band, 7) = "TDD" And Cells(Band, 8) <> "LTE" And Cells(Band, 9) = "NR" Then
                 .ForeColor.RGB = RGB(255, 0, 255) 'TDD mode, NR only color in Magenta
            ElseIf Cells(Band, 7) = "TDD" And Cells(Band, 8) = "LTE" And Cells(Band, 9) <> "NR" Then
                 .ForeColor.RGB = RGB(255, 255, 0) 'TDD mode, LTE only mode color in Yellow
            ElseIf Cells(Band, 7) = "TDD" And Cells(Band, 8) = "LTE" And Cells(Band, 9) = "NR" Then
                 .ForeColor.RGB = RGB(0, 0, 0) 'TDD mode, LTE and NR mode color in Black
            Else
                 .ForeColor.RGB = RGB(255, 255, 255) 'Non-TDD or Non-FDD mode color = no color
            End If
            
            'The following four codes are use to set the line similar is a flat end
            '.DashStyle = msoLineSquareDot
            '.DashStyle = msoLineSolid
            '.Style = msoLineThinThin
            .Style = msoLineSingle
            '.CapStyple = msoLineCapFlat
            
            .Weight = (ChartHeight - 200) / BandLast    'The line weight is proportional with the ChartHeight
            '.Weight = LineThinkness
            '.Style = msoLineCapSquare
            .Transparency = 0
        End With
        With ActiveChart.FullSeriesCollection(Band).Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0)
            .Transparency = 0
            .Solid
        End With
        'ActiveChart.SetElement (msoElementDataLabelCallout)
    End If
Next Band

Call PlotChart(UplinkChart, "Mobile Band vs Frequency", "Uplink Frequency (MHz)", UplinkBandTitle, 0)

End Sub
Sub Plot_LTE_Band_Downlink()
'
' Plot_NR ????X
'
'
Dim Band As Integer
Dim BandHighest As Integer
        
Dim cht As ChartObject
Dim ChartName As String

Call CreateChart(DownlinkChart)

For Band = BandFirst To BandLast
    ActiveSheet.ChartObjects(DownlinkChart).Activate
    ActiveChart.SeriesCollection.NewSeries
    If Worksheets(DataSource).Cells(Band, DownlinkMinFreq) <> "N/A" Then
        ActiveChart.FullSeriesCollection(Band).Name = Worksheets(DataSource).Cells(Band, BandNumber1)
        ActiveChart.FullSeriesCollection(Band).XValues = Worksheets(DataSource).Range(Cells(Band, DownlinkMinFreq), Cells(Band, DownlinkMaxFreq))
        ActiveChart.FullSeriesCollection(Band).Values = Worksheets(DataSource).Range(Cells(Band, BandNumber1), Cells(Band, BandNumber2))
        ActiveSheet.ChartObjects(DownlinkChart).Activate
        ActiveChart.FullSeriesCollection(Band).Select
        With Selection.Format.Line
            .Visible = msoTrue
            If Cells(Band, 7) = "FDD" And Cells(Band, 8) <> "LTE" And Cells(Band, 9) = "NR" Then
                 .ForeColor.RGB = RGB(0, 0, 255) 'FDD mode, NR only color in Blue
            ElseIf Cells(Band, 7) = "FDD" And Cells(Band, 8) = "LTE" And Cells(Band, 9) <> "NR" Then
                 .ForeColor.RGB = RGB(0, 255, 0) 'FDD mode, LTE only mode color in Green
            ElseIf Cells(Band, 7) = "FDD" And Cells(Band, 8) = "LTE" And Cells(Band, 9) = "NR" Then
                 .ForeColor.RGB = RGB(0, 255, 255) 'FDD mode, LTE and NR mode color in Aqua
            ElseIf Cells(Band, 7) = "TDD" And Cells(Band, 8) <> "LTE" And Cells(Band, 9) = "NR" Then
                 .ForeColor.RGB = RGB(255, 0, 255) 'TDD mode, NR only color in Magenta
            ElseIf Cells(Band, 7) = "TDD" And Cells(Band, 8) = "LTE" And Cells(Band, 9) <> "NR" Then
                 .ForeColor.RGB = RGB(255, 255, 0) 'TDD mode, LTE only mode color in Yellow
            ElseIf Cells(Band, 7) = "TDD" And Cells(Band, 8) = "LTE" And Cells(Band, 9) = "NR" Then
                 .ForeColor.RGB = RGB(0, 0, 0) 'TDD mode, LTE and NR mode color in Black
            Else
                 .ForeColor.RGB = RGB(255, 255, 255) 'Non-TDD or Non-FDD mode color = no color
            End If
            .DashStyle = msoLineSquareDot
            .DashStyle = msoLineSolid
            .Style = msoLineThinThin
            .Style = msoLineSingle
            .Weight = (ChartHeight - 200) / BandLast    'The line weight is proportional with the ChartHeight
            .Transparency = 0.75
        End With
        With Selection.Format.Fill
            .Visible = msoTrue
            .ForeColor.RGB = RGB(255, 0, 0) 'FDD mode RX color
            .Transparency = 0.75
            .Solid
        End With
        'ActiveChart.SetElement (msoElementDataLabelCallout)
    End If
Next Band

Call PlotChart(DownlinkChart, "Mobile Band vs Frequency", "Downlink Frequency (MHz)", DownlinkBandTitle, ChartWidth + ChartGap)

End Sub

Sub CreateChart(ChartName As String)
Dim cht As ChartObject
'Dim ChartName As String
'ChartName = UplinkChart
For Each cht In ActiveSheet.ChartObjects
    If cht.Name = ChartName Then
        cht.Delete
    End If
Next cht
    
' Create and define chart name, size
ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select
ActiveChart.Parent.Name = ChartName
With ActiveSheet.Shapes(ChartName)
    .Width = ChartWidth
    .Height = ChartHeight
End With
End Sub

Sub PlotChart(ChartName As String, MainTitle As String, XAxisTitle As String, YAxisTitle As String, ChartOffset As Integer)
'Define chart scale
    BandHighest = Cells(BandLast, 1)
    With ActiveChart.Axes(xlCategory)
        .MaximumScale = FreqMax
        .MinimumScale = FreqMin
        .MajorUnit = FreqStep
    End With
    With ActiveChart.Axes(xlValue)
        .MinimumScale = 0
        .MaximumScale = WorksheetFunction.Ceiling(BandHighest, 10)
        .MajorUnit = WorksheetFunction.Ceiling(BandHighest, 10) / (WorksheetFunction.Ceiling(BandHighest, 10) / BandStep)
        'ActiveChart.Axes(xlValue).MajorUnit = 10   'Set Major unit size is 10
    End With
    
    'Set the chart size and location
    With ActiveSheet.ChartObjects(ChartName)
        .Height = ChartHeight
        .Width = ChartWidth
        .Left = ChartLeft + ChartOffset
        .Top = ChartTop
    End With

'Update Charttitles
    ActiveSheet.ChartObjects(ChartName).Activate
    ActiveChart.SetElement (msoElementChartTitleAboveChart)
    With ActiveSheet.ChartObjects(ChartName).Chart.ChartTitle
        .Text = MainTitle
        .Format.TextFrame2.TextRange.Font.Size = 1.5 * WorksheetFunction.Min((ChartHeight / TitleRatio), (ChartWidth / TitleRatio))
    End With
    
'Update XAxis, YAxis titles
    ActiveSheet.ChartObjects(ChartName).Activate
    With ActiveChart
        With .Axes(xlCategory, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = XAxisTitle
            .AxisTitle.Font.Size = WorksheetFunction.Min((ChartHeight / TitleRatio), (ChartWidth / TitleRatio))
        End With
        With .Axes(xlValue, xlPrimary)
            .HasTitle = True
            .AxisTitle.Text = YAxisTitle
            .AxisTitle.Font.Size = WorksheetFunction.Min((ChartHeight / TitleRatio), (ChartWidth / TitleRatio))
            End With
    End With
    
    ActiveChart.SetElement (msoElementLegendNone)   'Hidden the legend
    'ActiveChart.SetElement (msoElementDataLabelCallout)
    ActiveWorkbook.Save

End Sub


Sub ¥¨¶°24()
'
' ¥¨¶°24 ¥¨¶°
'
ActiveChart.FullSeriesCollection(75).Select
    ActiveChart.SetElement (msoElementDataLabelCallout)     'Enable Data Label and text
    ActiveChart.FullSeriesCollection(75).Points(1).DataLabel.Select
    Selection.Left = 233.774
    Selection.Top = 40.98
    ActiveChart.FullSeriesCollection(75).Points(2).DataLabel.Select
    Selection.Left = 292.916
    Selection.Top = 53.398
    Debug.Print Selection.Position
'
End Sub
Sub CapTypeFlat()

Application.CommandBars("Format Object").Visible = True
ActiveChart.SeriesCollection(1).Select
    With Application
        .SendKeys "^1"                          ' format dialog
        .SendKeys "{DOWN}{DOWN}{DOWN}{DOWN}"    ' Line style
        .SendKeys "{TAB}{TAB}{TAB}{TAB}"        ' Cap type
        .SendKeys "F{RETURN}"                   ' Flat
        .SendKeys "{ESC}"                       ' close dialog
    End With

End Sub

