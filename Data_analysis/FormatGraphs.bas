Attribute VB_Name = "FormatGraphs"

Sub Format_eGraphs()

'compiled October 2024

'This sub was used to format graphs of experimental data (holdup, pressure drop and entrainment, at 45, 50 & 55 degrees C and 120 & 140 bars), so that all graphs could quickly be made to look consistent
'The code below is for plotting entrainment at 45, 50 & 55 dC only, where the chart has already been named ("Chart_e_45") and is on a named tab ("Summary45C")

Dim srs As Series
Dim chrt As Chart

Set chrt = Sheets("Summary45C").ChartObjects("Chart_e_45").Chart

'Colour guide (specified by university branding policy):
'   maroon - RGB(97, 34, 59)
'   Stellenbosch University gold - RGB(183, 153, 97)
'   Engineering gold - RGB(235, 169, 0)
'   grey - RGB(77, 83, 86)

'Format chart size
chrt.ChartArea.Height = 6.79 / 2.54 * 72 'to be ~7 cm
chrt.ChartArea.Width = 9.8 / 2.54 * 72 'to be ~10 cm
chrt.PlotArea.Height = 6.35 / 2.54 * 72
chrt.PlotArea.Width = 9.05 / 2.54 * 72
chrt.PlotArea.Left = 0.75 / 2.54 * 72

'Format x-axis and label
chrt.SetElement (msoElementPrimaryCategoryAxisShow) 'resets axis value formatting
chrt.Axes(xlCategory, xlPrimary).MinimumScale = 0.1
chrt.Axes(xlCategory, xlPrimary).MaximumScale = 2
chrt.Axes(xlCategory, xlPrimary).MinorUnit = 0.1
chrt.Axes(xlCategory, xlPrimary).MajorUnit = 1
chrt.Axes(xlCategory, xlPrimary).HasMinorGridlines = True
chrt.Axes(xlCategory, xlPrimary).MinorGridlines.Border.ColorIndex = 15 'grey vertical gridlines
chrt.Axes(xlCategory, xlPrimary).Border.ColorIndex = 1
chrt.Axes(xlCategory, xlPrimary).Crosses = xlMinimum
chrt.SetElement (msoElementPrimaryCategoryAxisLogScale) 'logarithmic scale on the x-axis, with base 10 (as set by the minimum scale value of 0.1 or 10^-1)
chrt.Axes(xlCategory, xlPrimary).TickLabels.Font.Name = "Trebuchet MS"
chrt.Axes(xlCategory, xlPrimary).TickLabels.Font.Size = 8
chrt.Axes(xlCategory, xlPrimary).AxisTitle.Top = 6.3 / 2.54 * 72
chrt.Axes(xlCategory, xlPrimary).AxisTitle.Left = 2.5 / 2.54 * 72
chrt.Axes(xlCategory, xlPrimary).AxisTitle.Text = "Superficial liquid velocity, uL (mm/s)"
chrt.Axes(xlCategory, xlPrimary).AxisTitle.Font.Name = "Trebuchet MS"
chrt.Axes(xlCategory, xlPrimary).AxisTitle.Font.Size = 10

'Format y-axis and label
chrt.SetElement (msoElementPrimaryValueAxis) 'resest axis value formatting
chrt.Axes(xlValue, xlPrimary).MinimumScale = 0
chrt.Axes(xlValue, xlPrimary).MaximumScale = 0.015
chrt.Axes(xlValue, xlPrimary).MinorUnit = 0.001
chrt.Axes(xlValue, xlPrimary).MajorUnit = 0.003
chrt.Axes(xlValue, xlPrimary).HasMinorGridlines = False
chrt.Axes(xlValue, xlPrimary).HasMajorGridlines = False
chrt.Axes(xlValue, xlPrimary).Border.ColorIndex = 1
chrt.Axes(xlValue, xlPrimary).Crosses = xlMinimum
chrt.Axes(xlValue, xlPrimary).TickLabels.Font.Name = "Trebuchet MS"
chrt.Axes(xlValue, xlPrimary).TickLabels.Font.Size = 8
chrt.Axes(xlValue, xlPrimary).AxisTitle.Left = 0 / 2.54 * 72
chrt.Axes(xlValue, xlPrimary).AxisTitle.Top = 1.5 / 2.54 * 72
chrt.Axes(xlValue, xlPrimary).AxisTitle.Text = "Entrainment ([g/g]"
chrt.Axes(xlValue, xlPrimary).AxisTitle.Font.Name = "Trebuchet MS"
chrt.Axes(xlValue, xlPrimary).AxisTitle.Font.Size = 10

'Format legend
chrt.Legend.Font.Name = "Trebuchet MS"
chrt.Legend.Font.Size = 9
chrt.Legend.Format.Fill.ForeColor.RGB = RGB(255, 255, 255) 'white solid fill
chrt.Legend.Format.Line.Visible = msoTrue
chrt.Legend.Format.Line.ForeColor.RGB = RGB(100, 100, 100) 'grey solid border
chrt.Legend.Left = 2 / 2.54 * 72
chrt.Legend.Width = 4 / 2.54 * 72
chrt.Legend.Height = 0.9 / 2.54 * 72
chrt.Legend.LegendEntries(1).LegendKey.Height = 0.2 / 2.54 * 72

'Format 55 dC data set: grey diamonds, semi-trasparent filling
Set srs = chrt.SeriesCollection(1)
srs.Name = "55°C"
srs.MarkerStyle = xlMarkerStyleDiamond
srs.Format.Line.Weight = 1.5
srs.Format.Line.Visible = msoTrue
srs.Format.Line.Visible = msoTriStateToggle
srs.Format.Line.ForeColor.RGB = RGB(255, 255, 255)
srs.Format.Fill.BackColor.RGB = RGB(77, 83, 86)
srs.MarkerForegroundColor = RGB(77, 83, 86)
srs.Format.Fill.ForeColor.RGB = RGB(77, 83, 86)
srs.Format.Fill.Solid
srs.Format.Fill.Transparency = 0.5
srs.MarkerSize = 6

'Format 50 dC data set: maroon circles
Set srs = chrt.SeriesCollection(2)
srs.Name = "50°C"
srs.MarkerStyle = xlMarkerStyleCircle
srs.MarkerForegroundColor = RGB(97, 34, 59)
srs.MarkerBackgroundColor = RGB(97, 34, 59)
srs.MarkerSize = 6

'Format 45 dC data set: Engineering gold squares with maroon border
Set srs = chrt.SeriesCollection(3)
srs.Name = "45°C"
srs.MarkerStyle = xlMarkerStyleSquare
srs.MarkerForegroundColor = RGB(97, 34, 59)
srs.MarkerBackgroundColor = RGB(235, 169, 0)
srs.MarkerSize = 6

End Sub
