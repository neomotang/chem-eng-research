Attribute VB_Name = "f8CopyClearHome"

Sub CopyClearHome()

'compiled August 2024

'This sub copies the results stored in row 21 onwards on the "Home" tab, and saves it on the "Results" tab (Sheet3).
'After copying, the "Home" tab is then cleared.
    
Application.ScreenUpdating = False
Application.DisplayAlerts = False

Sheet1.Activate
DataCount = Sheet1.Range("A20", Range("A20").End(xlDown)).Rows.Count
Application.CutCopyMode = False

If DataCount <= 1 Or DataCount = 1048557 Then '1 = no entries, i.e. just heading row contains text; 1048557 = no entries but Excel has also selected all the blank rows til bottom of spreadsheet
    Sheet1.Activate
    Range("F8", Range("F10")).Value = ""
    Range("A21").Select
    Exit Sub
Else
    If DataCount = 2 Then '2 = only a single entry
        Range("A21", Range("AI21")).Select
        Selection.Copy
        Sheet3.Activate
        DataEnd = Sheet3.Range("A1", Range("A1").End(xlDown)).Rows.Count
        Sheet3.Cells(DataEnd + 1, 1).Select
        ActiveSheet.Paste
    Else
        Range("A21", Range("AI21").End(xlDown)).Select
        Selection.Copy
        Sheet3.Activate
        DataEnd = Sheet3.Range("A1", Range("A1").End(xlDown)).Rows.Count
        Sheet3.Cells(DataEnd + 1, 1).Select
        ActiveSheet.Paste
    End If
    
    Application.CutCopyMode = False
        
End If
        
Sheet1.Activate
Range("F8", Range("F10")).Value = ""
Range("A21", Range("AI21").End(xlDown)).Select
Selection.Value = ""

Range("A21").Select
Application.ScreenUpdating = True
Application.DisplayAlerts = True

End Sub

