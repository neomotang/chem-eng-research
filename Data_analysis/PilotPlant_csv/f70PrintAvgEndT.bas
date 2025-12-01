Attribute VB_Name = "f70PrintAvgEndT"

Sub PrintEndTemps(ByVal FCount As Integer)

'compiled August 2024

'This sub stores the averages of the data in PrTemp.csv in columns U-AB and the appropriate row on the "Home" tab.
'The required input is the row for the current experimental run under consideration.

'Avg for T1-1
Sheet1.Cells(FCount, 21).Value = avgT11
Sheet1.Cells(FCount, 21).NumberFormat = "0.0"

'Avg for T2-1
Sheet1.Cells(FCount, 22).Value = avgT21
Sheet1.Cells(FCount, 22).NumberFormat = "0.0"
                
'Avg for T3-1
Sheet1.Cells(FCount, 23).Value = avgT31
Sheet1.Cells(FCount, 23).NumberFormat = "0.0"

'Avg & stdv for T3-2
Sheet1.Cells(FCount, 24).Value = avgT32
Sheet1.Cells(FCount, 24).NumberFormat = "0.0"
Sheet1.Cells(FCount, 25).Value = stdT32

'Avg for T3-3
Sheet1.Cells(FCount, 26).Value = avgT33
Sheet1.Cells(FCount, 26).NumberFormat = "0.0"
                
'Avg for T4-1
            
Sheet1.Cells(FCount, 27).Value = avgT41
Sheet1.Cells(FCount, 27).NumberFormat = "0.0"

'Avg for T4-2
Sheet1.Cells(FCount, 28).Value = avgT42
Sheet1.Cells(FCount, 28).NumberFormat = "0.0"

End Sub
