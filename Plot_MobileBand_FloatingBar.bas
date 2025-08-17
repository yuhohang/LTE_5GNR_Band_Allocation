Attribute VB_Name = "Plot_MobileBand_FloatingBar"
Const BandFirst As Integer = 1, BandLast As Integer = 82 'Avoid larger than 66
Const BandNumber As Integer = 1
Const UplinkMinFreq As Integer = 3, UplinkMaxFreq As Integer = 4
Const DownlinkMinFreq As Integer = 5, DownlinkMaxFreq As Integer = 6
Const FreqMin As Integer = 0, FreqMax As Integer = 6000, FreqStep As Integer = 500
Const BandStep As Integer = 5
Const UplinkChart As String = "Chart 1", DownlinkChart As String = "Chart 2"
Const UplinkBandTitle As String = "LTE & NR Band"
Const DownlinkBandTitle As String = "LTE & NR Band"
Const DataSource As String = "LTE_NR_Bar"   'Define the chart source worksheet name
Const ChartHeight As Integer = 1000, ChartWidth As Integer = 300, ChartTop As Integer = 0, ChartLeft As Integer = 500

Sub LTE_5GNR_Uplink_FloatingBar_Creation()
Attribute LTE_5GNR_Uplink_FloatingBar_Creation.VB_ProcData.VB_Invoke_Func = " \n14"
'
' ⅷ떠6 ⅷ떠
'

'
    Dim rngTemp As Range
    ActiveSheet.Shapes.AddChart2(297, xlBarStacked).Select
    ActiveChart.Parent.Name = UplinkChart
    ActiveChart.SetSourceData Source:=Range( _
        "LTE_NR_Bar!$B$2:$B$82,LTE_NR_Bar!$C$2:$C$82,LTE_NR_Bar!$E$2:$E$82")
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(1).Select          'Hidden Start Freq
    Selection.Format.Fill.Visible = msoFalse
    Selection.Format.Line.Visible = msoFalse
    ActiveChart.PlotArea.Select
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(255, 0, 0)
        .Transparency = 0
        .Solid
    End With
    ActiveChart.ChartGroups(1).GapWidth = 40
    
    'Define chart scale
    BandHighest = Cells(BandLast, 1)
    ActiveSheet.ChartObjects(UplinkChart).Activate
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MaximumScale = FreqMax
    ActiveChart.Axes(xlValue).MinimumScale = FreqMin
    ActiveChart.Axes(xlValue).MajorUnit = FreqStep
    
    ActiveSheet.ChartObjects(UplinkChart).Height = ChartHeight
    ActiveSheet.ChartObjects(UplinkChart).Width = ChartWidth
    ActiveSheet.ChartObjects(UplinkChart).Left = ChartLeft
    ActiveSheet.ChartObjects(UplinkChart).Top = ChartTop
    
    'Add chart title
    With ActiveChart
      With .Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = UplinkBandTitle
      End With
      With .Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Uplink Frequency (MHz)"
      End With
    End With
    
    ActiveWorkbook.Save
End Sub
Sub LTE_5GNR_Downlink_FloatingBar_Creation()
'
' ⅷ떠6 ⅷ떠
'

'
    ActiveSheet.Shapes.AddChart2(297, xlBarStacked).Select
    ActiveChart.Parent.Name = DownlinkChart
    ActiveChart.SetSourceData Source:=Range( _
        "LTE_NR_Bar!$B$2:$B$82,LTE_NR_Bar!$G$2:$G$82,LTE_NR_Bar!$I$2:$I$82")
    ActiveChart.ChartArea.Select
    ActiveChart.FullSeriesCollection(1).Select
    Selection.Format.Fill.Visible = msoFalse
    Selection.Format.Line.Visible = msoFalse
    
    ActiveChart.PlotArea.Select
    ActiveChart.FullSeriesCollection(2).Select
    With Selection.Format.Fill
        .Visible = msoTrue
        .ForeColor.RGB = RGB(0, 255, 0)
        .Transparency = 0
        .Solid
    End With
    ActiveChart.ChartGroups(1).GapWidth = 40
    
    'Define chart scale
    BandHighest = Cells(BandLast, 1)
    ActiveSheet.ChartObjects(DownlinkChart).Activate
    ActiveChart.Axes(xlValue).Select
    ActiveChart.Axes(xlValue).MaximumScale = FreqMax
    ActiveChart.Axes(xlValue).MinimumScale = FreqMin
    ActiveChart.Axes(xlValue).MajorUnit = FreqStep
    
    ActiveSheet.ChartObjects(DownlinkChart).Height = ChartHeight
    ActiveSheet.ChartObjects(DownlinkChart).Width = ChartWidth
    ActiveSheet.ChartObjects(DownlinkChart).Left = ChartLeft + ChartWidth + 6
    ActiveSheet.ChartObjects(DownlinkChart).Top = ChartTop
    
     With ActiveChart
      With .Axes(xlCategory, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = DownlinkBandTitle
      End With
      With .Axes(xlValue, xlPrimary)
        .HasTitle = True
        .AxisTitle.Text = "Downlink Frequency (MHz)"
      End With
    End With
   
    ActiveWorkbook.Save
End Sub
