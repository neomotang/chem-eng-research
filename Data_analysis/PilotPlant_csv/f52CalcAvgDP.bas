Attribute VB_Name = "f52CalcAvgDP"

Public avgDP31, stdDP31 As Double

Sub GetAvgDP(ByVal RCount As Long)

'compiled August 2024

'This sub calculates the averages of the data in PrDp.csv. The required input is the row identified as the approximate time (either start or end).
'Calculation of averages are over a time span of about 2 minutes. Results are stored in public variables.

Sheet2.Activate

'Avg & stdv for DP
Range("AE" & RCount - 559, Range("AE" & RCount)).Select
Selection.Copy
Range("AG3").Select
ActiveSheet.Paste
Sheet2.Range("AG1").Value = "=AVERAGE(AG3:AG562)"
avgDP31 = Sheet2.Range("AG1").Value
Sheet2.Range("AG2").Value = "=STDEV.S(AG3:AG562)"
stdDP31 = Sheet2.Range("AG2").Value

End Sub
