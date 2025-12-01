Attribute VB_Name = "f71PrintAvgEndF"

Sub PrintEndFlows(ByVal FCount As Integer)

'compiled August 2024

'This sub stores the averages of the data in PrFlow.csv in columns AC-AF and the appropriate row on the "Home" tab.
'The required input is the row for the current experimental run under consideration.

'Avg for DP
Sheet1.Cells(FCount, 29).Value = avgDP
Sheet1.Cells(FCount, 29).NumberFormat = "0.00"

'Avg for Flow
Sheet1.Cells(FCount, 30).Value = avgFlow
Sheet1.Cells(FCount, 30).NumberFormat = "0.0"
                
'Avg for P4-1
Sheet1.Cells(FCount, 31).Value = avgP41
Sheet1.Cells(FCount, 31).NumberFormat = "0.0"

'Avg for P4-2
Sheet1.Cells(FCount, 32).Value = avgP42
Sheet1.Cells(FCount, 32).NumberFormat = "0.0"

End Sub
