Attribute VB_Name = "f51CalcAvgF"

Public avgDP, avgFlow, avgP41, avgP42 As Double

Sub GetAvgFlows(ByVal RCount As Integer)

'compiled August 2024

'This sub calculates the averages of the data in PrFlow.csv. The required input is the row identified as the approximate time (either start or end).
'Calculation of averages are over a time span of about 5 minutes. Results are stored in public variables.

Sheet2.Activate

'Avg for DP
Range("N" & RCount - 29, Range("N" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgDP = Sheet2.Range("X1").Value

'Avg for Flow
Range("O" & RCount - 29, Range("O" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgFlow = Sheet2.Range("X1").Value

'Avg for P4-1
Range("P" & RCount - 29, Range("P" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgP41 = Sheet2.Range("X1").Value

'Avg for P4-2
Range("Q" & RCount - 29, Range("Q" & RCount)).Select
Selection.Copy
Range("X3").Select
ActiveSheet.Paste
Sheet2.Range("X1").Value = "=AVERAGE(X3:X32)"
avgP42 = Sheet2.Range("X1").Value

End Sub

