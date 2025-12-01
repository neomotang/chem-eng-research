Attribute VB_Name = "f60PrintAvgStartT"

Sub PrintStartTemps(ByVal FCount As Integer)

'compiled August 2024

'This sub stores the averages of the data in PrTemp.csv in columns E-L and the appropriate row on the "Home" tab.
'The required input is the row for the current experimental run under consideration.

'Avg for T1-1
Sheet1.Cells(FCount, 5).Value = avgT11
Sheet1.Cells(FCount, 5).NumberFormat = "0.0"

'Avg for T2-1
Sheet1.Cells(FCount, 6).Value = avgT21
Sheet1.Cells(FCount, 6).NumberFormat = "0.0"
                
'Avg for T3-1
Sheet1.Cells(FCount, 7).Value = avgT31
Sheet1.Cells(FCount, 7).NumberFormat = "0.0"

'Avg & stdv for T3-2
Sheet1.Cells(FCount, 8).Value = avgT32
Sheet1.Cells(FCount, 8).NumberFormat = "0.0"
Sheet1.Cells(FCount, 9).Value = stdT32

'Avg for T3-3
Sheet1.Cells(FCount, 10).Value = avgT33
Sheet1.Cells(FCount, 10).NumberFormat = "0.0"
                
'Avg for T4-1
Sheet1.Cells(FCount, 11).Value = avgT41
Sheet1.Cells(FCount, 11).NumberFormat = "0.0"

'Avg for T4-2
Sheet1.Cells(FCount, 12).Value = avgT42
Sheet1.Cells(FCount, 12).NumberFormat = "0.0"

End Sub

