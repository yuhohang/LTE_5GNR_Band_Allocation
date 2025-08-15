Attribute VB_Name = "Plot_Band"
Const BandFirst As Integer = 1
Const BandLast As Integer = 64 'Avoid larger than 66
Const BandNumber1 As Integer = 1
Const BandNumber2 As Integer = 2
Const UplinkMinFreq As Integer = 3
Const UplinkMaxFreq As Integer = 4
Const DownlinkMinFreq As Integer = 5
Const DownlinkMaxFreq As Integer = 6
Const FreqMax As Integer = 6000
Const FreqMin As Integer = 0
Const FreqStep As Integer = 1000
Const BandStep As Integer = 5
Const UplinkChart As String = "Chart 1"
Const DownlinkChart As String = "Chart 2"
Const LineThinkness As Integer = 4
Const UplinkBandTitle As String = "LTE Band"
Const DownlinkBandTitle As String = "LTE Band"



Sub Plot_LTE_Band_Uplink()
'
' Plot_NR ????X
'
'
Dim Band As Integer
Dim BandHighest As Integer
       
Dim cht As ChartObject
Dim chartName As String
    chartName = UplinkChart
    For Each cht In ActiveSheet.ChartObjects
      If cht.Name = chartName Then
        cht.Delete
        Exit Sub ' Exit after deleting the chart
      End If
    Next cht
    
' Create and define chart name, size
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select
    ActiveChart.Parent.Name = UplinkChart
    ActiveSheet.Shapes(UplinkChart).Width = 425.1968503937
    ActiveSheet.Shapes(UplinkChart).Height = 708.6614173228
       
For Band = BandFirst To BandLast
    ActiveSheet.ChartObjects(UplinkChart).Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(Band).Name = Worksheets("NR").Cells(Band, BandNumber1)
    ActiveChart.FullSeriesCollection(Band).XValues = Worksheets("NR").Range(Cells(Band, UplinkMinFreq), Cells(Band, UplinkMaxFreq))
    ActiveChart.FullSeriesCollection(Band).Values = Worksheets("NR").Range(Cells(Band, BandNumber1), Cells(Band, BandNumber2))
    'ActiveChart.FullSeriesCollection(002).Name = "=NR!$A$002"
    'ActiveChart.FullSeriesCollection(002).XValues = "=NR!$B$002,NR!$C$002"
    'ActiveChart.FullSeriesCollection(002).Values = "=NR!$A$002,NR!$A$002"
    'ActiveSheet.ChartObjects(UplinkChart).Activate
    ActiveChart.FullSeriesCollection(Band).Select
    'ActiveChart.FullSeriesCollection(001).Select
    With Selection.Format.Line
        .Visible = msoTrue
        If Cells(Band, 7) = "FDD" Then
             .ForeColor.RGB = RGB(255, 0, 0) 'FDD mode TX color in Red
        ElseIf Cells(Band, 7) = "TDD" Then
             .ForeColor.RGB = RGB(0, 0, 255) 'TDD mode TX color in Blue
        Else
             .ForeColor.RGB = RGB(255, 255, 255) 'Non-TDD or FDD mode color = no color
        End If
        .DashStyle = msoLineSquareDot
        .DashStyle = msoLineSolid
        .Style = msoLineThinThin
        .Style = msoLineSingle
        .Weight = LineThinkness
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
    'ActiveChart.Axes(xlValue).MajorUnit = 10
    
    With ActiveChart
      With .Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Frequency (MHz)"
      End With
      With .Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = UplinkBandTitle
      End With
    End With

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
        Exit Sub ' Exit after deleting the chart
      End If
    Next cht

' Create and define chart name, size
    ActiveSheet.Shapes.AddChart2(240, xlXYScatterLinesNoMarkers).Select
    ActiveChart.Parent.Name = DownlinkChart
    ActiveSheet.Shapes(DownlinkChart).Width = 425.1968503937
    ActiveSheet.Shapes(DownlinkChart).Height = 708.6614173228
        
For Band = BandFirst To BandLast
    ActiveSheet.ChartObjects(DownlinkChart).Activate
    ActiveChart.SeriesCollection.NewSeries
    ActiveChart.FullSeriesCollection(Band).Name = Worksheets("NR").Cells(Band, BandNumber1)
    ActiveChart.FullSeriesCollection(Band).XValues = Worksheets("NR").Range(Cells(Band, DownlinkMinFreq), Cells(Band, DownlinkMaxFreq))
    ActiveChart.FullSeriesCollection(Band).Values = Worksheets("NR").Range(Cells(Band, BandNumber1), Cells(Band, BandNumber2))
    ActiveSheet.ChartObjects(DownlinkChart).Activate
    ActiveChart.FullSeriesCollection(Band).Select
    With Selection.Format.Line
        .Visible = msoTrue
        If Cells(Band, 7) = "FDD" Then
             .ForeColor.RGB = RGB(0, 255, 0) 'FDD mode RX color in Green
        ElseIf Cells(Band, 7) = "TDD" Then
             .ForeColor.RGB = RGB(0, 0, 255) 'TDD mode TX color in Blue
        Else
             .ForeColor.RGB = RGB(255, 255, 255) 'Non-TDD or FDD mode color = no color
        End If
        .DashStyle = msoLineSquareDot
        .DashStyle = msoLineSolid
        .Style = msoLineThinThin
        .Style = msoLineSingle
        .Weight = LineThinkness
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
    'ActiveChart.Axes(xlValue).MajorUnit = 10
    
    With ActiveChart
      With .Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Frequency (MHz)"
      End With
      With .Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = DownlinkBandTitle
      End With
    End With

End Sub

