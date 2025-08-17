Attribute VB_Name = "Plot_MobileBand_XYChart"
Const BandStep As Integer = 5, BandFirst As Integer = 1, BandLast As Integer = 80 'Avoid larger than 66',
Const BandNumber1 As Integer = 1, BandNumber2 As Integer = 2
Const UplinkMinFreq As Integer = 3, UplinkMaxFreq As Integer = 4
Const DownlinkMinFreq As Integer = 5, DownlinkMaxFreq As Integer = 6
Const FreqStep As Integer = 500, FreqMin As Integer = 0, FreqMax As Integer = 6000
Const UplinkChart As String = "Chart 1", DownlinkChart As String = "Chart 2"
Const LineThinkness As Integer = 4
Const UplinkBandTitle As String = "LTE & NR Band", DownlinkBandTitle As String = "LTE & NR Band"
Const DataSource As String = "LTE_NR"   'Define the chart source worksheet name
Const ChartHeight As Integer = 5 * 150, ChartWidth As Integer = (4 + 1) * 120
Const ChartTop As Integer = 0, ChartLeft As Integer = 500
'msoLineCapFlat
'.SendKeys "F{RETURN}{ENTER}" ' Flat

Sub Plot_LTE_Band_Uplink()
'
' Plot_NR ????X
'
'
Dim Band As Integer
Dim BandHighest As Integer
      
'Delete the Uplink chart
Dim cht As ChartObject
Dim chartName As String
    chartName = UplinkChart
    For Each cht In ActiveSheet.ChartObjects
      If cht.Name = chartName Then
        cht.Delete
        'Exit Sub ' Exit after deleting the chart
      End If
    Next cht
    
' Create and define chart name, size
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select
    ActiveChart.Parent.Name = UplinkChart
    ActiveSheet.Shapes(UplinkChart).Width = ChartWidth
    ActiveSheet.Shapes(UplinkChart).Height = ChartHeight
      
'Update the serie data of the chart source
For Band = BandFirst To BandLast
    ActiveSheet.ChartObjects(UplinkChart).Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(Band).Name = Worksheets(DataSource).Cells(Band, BandNumber1)
    ActiveChart.FullSeriesCollection(Band).XValues = Worksheets(DataSource).Range(Cells(Band, UplinkMinFreq), Cells(Band, UplinkMaxFreq))
    ActiveChart.FullSeriesCollection(Band).Values = Worksheets(DataSource).Range(Cells(Band, BandNumber1), Cells(Band, BandNumber2))
    'ActiveSheet.ChartObjects(UplinkChart).Activate
    'Update the chart line format and color
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
        'The following four codes are use to set the line similar is a flat end
        .DashStyle = msoLineSquareDot
        .DashStyle = msoLineSolid
        .Style = msoLineThinThin
        .Style = msoLineSingle
        .Weight = (ChartHeight - 200) / BandLast    'The line weight is proportional with the ChartHeight
        '.Weight = LineThinkness
        '.Style = msoLineCapSquare
        .Transparency = 0
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
        .Solid
    End With
    'ActiveChart.SetElement (msoElementDataLabelCallout)
Next Band

'Define chart scale
    BandHighest = Cells(BandLast, 1)
    ActiveSheet.ChartObjects(UplinkChart).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MaximumScale = FreqMax
    ActiveChart.Axes(xlCategory).MinimumScale = FreqMin
    ActiveChart.Axes(xlCategory).MajorUnit = FreqStep
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 0
    ActiveChart.Axes(xlValue).MaximumScale = WorksheetFunction.Ceiling(BandHighest, 10)
    ActiveChart.Axes(xlValue).MajorUnit = WorksheetFunction.Ceiling(BandHighest, 10) / (WorksheetFunction.Ceiling(BandHighest, 10) / BandStep)
    'ActiveChart.Axes(xlValue).MajorUnit = 10   'Set Major unit size is 10
    
    'Set the chart size and location
    ActiveSheet.ChartObjects(UplinkChart).Height = ChartHeight
    ActiveSheet.ChartObjects(UplinkChart).Width = ChartWidth
    ActiveSheet.ChartObjects(UplinkChart).Left = ChartLeft
    ActiveSheet.ChartObjects(UplinkChart).Top = ChartTop
    
'Add chart title
    With ActiveChart
      With .Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Uplink Frequency (MHz)"
      End With
      With .Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = UplinkBandTitle
      End With
    End With
    
    ActiveChart.SetElement (msoElementLegendNone)   'Hidden the legend
    ActiveWorkbook.Save

End Sub

Sub Plot_LTE_Band_Downlink()
'
' Plot_NR ????X
'
'
Dim Band As Integer
Dim BandHighest As Integer
        
Dim cht As ChartObject
Dim chartName As String
    chartName = DownlinkChart
    For Each cht In ActiveSheet.ChartObjects
      If cht.Name = chartName Then
        cht.Delete
        'Exit Sub ' Exit after deleting the chart
      End If
    Next cht

' Create and define chart name, size
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select
    ActiveChart.Parent.Name = DownlinkChart
    ActiveSheet.Shapes(DownlinkChart).Width = ChartWidth
    ActiveSheet.Shapes(DownlinkChart).Height = ChartHeight
Debug.Print "Line weight = "; (ChartHeight - 200) / BandLast
For Band = BandFirst To BandLast
    ActiveSheet.ChartObjects(DownlinkChart).Activate
    ActiveChart.SeriesCollection.NewSeries
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
        .Transparency = 0
    End With
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0) 'FDD mode RX color
        .Transparency = 0
        .Solid
    End With
    'ActiveChart.SetElement (msoElementDataLabelCallout)
Next Band

    BandHighest = Cells(BandLast, 1)
    ActiveSheet.ChartObjects(DownlinkChart).Activate
    ActiveChart.Axes(xlCategory).Select
    ActiveChart.Axes(xlCategory).MaximumScale = FreqMax
    ActiveChart.Axes(xlCategory).MinimumScale = FreqMin
    ActiveChart.Axes(xlCategory).MajorUnit = FreqStep
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MinimumScale = 0
    ActiveChart.Axes(xlValue).MaximumScale = WorksheetFunction.Ceiling(BandHighest, 10)
    ActiveChart.Axes(xlValue).MajorUnit = WorksheetFunction.Ceiling(BandHighest, 10) / (WorksheetFunction.Ceiling(BandHighest, 10) / BandStep)
    'ActiveChart.Axes(xlValue).MajorUnit = 10   'Set Major unit size is 10
    
    'Set the chart size and location
    ActiveSheet.ChartObjects(DownlinkChart).Height = ChartHeight
    ActiveSheet.ChartObjects(DownlinkChart).Width = ChartWidth
    ActiveSheet.ChartObjects(DownlinkChart).Left = ChartLeft + ChartWidth + 5
    ActiveSheet.ChartObjects(DownlinkChart).Top = ChartTop
    
    With ActiveChart
      With .Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Downlink Frequency (MHz)"
      End With
      With .Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = DownlinkBandTitle
      End With
    End With
    
    ActiveChart.SetElement (msoElementLegendNone)   'Hidden the Legend
    ActiveWorkbook.Save

End Sub

Sub CapTypeFlat()

 Application.SendKeys "F{RETURN}{ENTER}" ' Flat

End Sub
