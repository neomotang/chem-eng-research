Attribute VB_Name = "f72PrintAvgEndDP"

Sub PrintEndDP(ByVal FCount As Integer)

'compiled August 2024

'This sub stores the averages of the data in PrDp.csv in columns AH-AI and the appropriate row on the "Home" tab.
'The required input is the row for the current experimental run under consideration.

'Avg for DP
Sheet1.Cells(FCount, 34).Value = avgDP31
Sheet1.Cells(FCount, 34).NumberFormat = "0.00"
Sheet1.Cells(FCount, 35).Value = stdDP31

End Sub
