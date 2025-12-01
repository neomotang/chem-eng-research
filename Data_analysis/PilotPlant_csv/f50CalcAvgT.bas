Attribute VB_Name = "f50CalcAvgT"

Public avgT11, avgT21, avgT31, avgT32, stdT32, avgT33, avgT41, avgT42 As Double

Sub GetAvgTemps(ByVal RCount As Integer)

'compiled August 2024

'This sub calculates the averages of the data in PrTemp.csv. The required input is the row identified as the approximate time (either start or end).
'Calculation of averages are over a time span of about 5 minutes. Results are stored in public variables.

Sheet2.Activate

'Avg for T1-1
Range("E" & RCount - 29, Range("E" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgT11 = Sheet2.Range("X1").Value

'Avg for T2-1
Range("F" & RCount - 29, Range("F" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgT21 = Sheet2.Range("X1").Value

'Avg for T3-1
Range("G" & RCount - 29, Range("G" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgT31 = Sheet2.Range("X1").Value

'Avg & stdv for T3-2
Range("H" & RCount - 29, Range("H" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgT32 = Sheet2.Range("X1").Value
Sheet2.Range("X2").Value = "=STDEV.S(X3:X32)"
stdT32 = Sheet2.Range("X2").Value

'Avg for T3-3
Range("I" & RCount - 29, Range("I" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgT33 = Sheet2.Range("X1").Value

'Avg for T4-1
Range("J" & RCount - 29, Range("J" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgT41 = Sheet2.Range("X1").Value

'Avg for T4-2
Range("K" & RCount - 29, Range("K" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgT42 = Sheet2.Range("X1").Value

End Sub

