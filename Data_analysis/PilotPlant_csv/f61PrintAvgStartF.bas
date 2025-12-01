Attribute VB_Name = "f61PrintAvgStartF"

Sub PrintStartFlows(ByVal FCount As Integer)

'compiled August 2024

'This sub stores the averages of the data in PrFlow.csv in columns M-P and the appropriate row on the "Home" tab.
'The required input is the row for the current experimental run under consideration.

'Avg for DP
Sheet1.Cells(FCount, 13).Value = avgDP
Sheet1.Cells(FCount, 13).NumberFormat = "0.00"

'Avg for Flow
Sheet1.Cells(FCount, 14).Value = avgFlow
Sheet1.Cells(FCount, 14).NumberFormat = "0.0"
                
'Avg for P4-1
Sheet1.Cells(FCount, 15).Value = avgP41
Sheet1.Cells(FCount, 15).NumberFormat = "0.0"

'Avg for P4-2
Sheet1.Cells(FCount, 16).Value = avgP42
Sheet1.Cells(FCount, 16).NumberFormat = "0.0"

End Sub
