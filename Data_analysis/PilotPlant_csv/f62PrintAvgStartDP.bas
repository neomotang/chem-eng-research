Attribute VB_Name = "f62PrintAvgStartDP"

Sub PrintStartDP(ByVal FCount As Integer)

'compiled August 2024

'This sub stores the averages of the data in PrDp.csv in columns R-S and the appropriate row on the "Home" tab.
'The required input is the row for the current experimental run under consideration.

'Avg & stdv for DP
Sheet1.Cells(FCount, 18).Value = avgDP31
Sheet1.Cells(FCount, 18).NumberFormat = "0.00"
Sheet1.Cells(FCount, 19).Value = stdDP31

End Sub
