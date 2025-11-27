Attribute VB_Name = "MinFrequency"

'Option Explicit

Sub FindMinFreq()

'compiled August 2020

'To be used with GetFolderPathNFileNum.bas and LoopThroughFolder.bas

'This sub is called once the frequeny and "Data" heading have been copied to column L, and the voltage and "collection started at..." have been copied to column M, on Sheet3 ("Worksheet"). The minimum frequency and relevant stats will be determined for each block of data.

'Data in the .csv file is the frequency and voltage, with a note at the beginning of each data block indicating the date and time at which data collection started: "Data collection started at: Thu Jul 04 14:31:54 2019"
'Multiple periods of data collection may be in one file, but all data will be for the same chemical system, composition & temperature

'Frequency min.xlsm is the main file where the output will be recorded; it contains the following three tabs:
'   Sheet1 ("Home") - starting tab with buttons for code execution: a button executes sub "GetFolderPathNFileCount" (output as folder path and number of files in cells F8:F9 on "Home" tab), and another button executes sub "LoopThroughFold" (output as stats on "Results" tab)
'   Sheet3 ("Worksheet") - working space used to copy data from the .csv files
'   Sheet5 ("Results") - records stats from the quadratic regression and estimate of frequency; has the following columns: Filename (A), Min Freq [Hz] (B), FreqStep [Hz] (C), R^2 (D), A (x2 coefficient) (F), B (x coefficient) (G), C (y intercept) (H), Standard error of A (I), Standard error of B (J), Standard error of C (K), Standard error of Y estimate (L), F statistic (M), degrees of freedom (N), Sum of squares of the regression (O), and Sum of squares of the residuals (P)

    
    Dim DataBlock As Range
    
    Sheet3.Activate 'necessary so that 'Set DataBlock = Range("A1")' can work
    BeginCellX = "$L$1"
    BeginCellY = "$M$1"
    FoundAt = "$L$1" 'initialize to represent the first position of "Data"; might need to be updated in the If statement just below
    Set DataBlock = Range("$L$1") 'initialize to represent the first position of "Data"; might need to be updated in the If statement just below
    EndOfCol = "No"
    NumEntries = 0
    RowCount = 0
    
Do
    Sheet3.Activate
    Sheet3.Range(BeginCellX).Activate
    
    SearchTop = Range(BeginCellX).Offset(0, 0).Address 'assumes BeginCellX contains headers, i.e. "Data", and moves to the next row, which should have numbers
    SearchBottom = Range(BeginCellX).End(xlDown).Address 'finds the end of the column
    
    RowCount = Sheet3.Range(SearchTop, SearchBottom).Rows.count
    If RowCount = 0 Then 'meaning the sheet is blank
        Sheet5.Activate
        Range("A1").End(xlDown).Select 'bug over here that will go to the last possible row (A1048576) if there is only one row of text initially; hence adding units of the variables in the second row
        ActiveCell.Offset(1, 1).Value = "Sheet is blank" 'text entered in cell e.g. B4
        ActiveCell.Offset(1, 0).Value = Sheets(Sheets.count).Name 'the file name of the .csv file (which is currently the name of the data tab) is entered in cell e.g. A4
        EndOfCol = "Yes"
        Else
            Set DataBlock = Range(SearchTop, SearchBottom).Find("Data") 'use to find next instance of "Data" in column L
            
            If DataBlock.Address = FoundAt Then 'true if another instance of "Data" has not been found, meaning last block of numbers is the current
                If Range(FoundAt).Offset(3, 0) = "" Then 'true if "Data" happens to be the last entry with no numbers following; exit the loop; set to 3 instead of 1 so that at least 3 data points can be passed to the quadratic function
                    EndOfCol = "Yes"
                    GoTo EndOfLoop:
                    Else
                    Set DataBlock = Range(BeginCellX).End(xlDown).Offset(1, 0) 'move to end of column
                    EndOfCol = "Yes"
                    End If
                End If
        
        Sheet3.Range(FoundAt).Select
        BeginCellX = ActiveCell.Offset(1, 0).Address 'goes to the beginning of the numerical data block, e.g. L2
        BeginCellY = ActiveCell.Offset(1, 1).Address 'goes to the second column in the block, e.g. M2
            
        FoundAt = DataBlock.Address
        Sheet3.Range(FoundAt).Activate
        EndCellX = ActiveCell.Offset(-1, 0).Address
        EndCellY = ActiveCell.Offset(-1, 1).Address
        
        NumEntries = Sheet3.Range(BeginCellX, EndCellX).Rows.count
        
        Do While NumEntries < 3 And EndOfCol = "No"
            BeginCellX = FoundAt 'reintialize before searching for the next instance of Data
            SearchTop = Range(BeginCellX).Offset(0, 0).Address 'assumes BeginCellX contains headers, i.e. "Data", and moves to the next row, which should have numbers
            SearchBottom = Range(BeginCellX).End(xlDown).Address 'finds the end of the column
    
            Set DataBlock = Range(SearchTop, SearchBottom).Find("Data") 'use to find next instance of "Data" in column A
            
            If DataBlock.Address = FoundAt Then 'true if another instance of "Data" has not been found, meaning last block of numbers is the current
                Set DataBlock = Range(BeginCellX).End(xlDown).Offset(1, 0) 'move to end of column
                EndOfCol = "Yes"
                End If
            
            Sheet3.Range(FoundAt).Select
            BeginCellX = ActiveCell.Offset(1, 0).Address
            BeginCellY = ActiveCell.Offset(1, 1).Address
            
            FoundAt = DataBlock.Address
            Sheet3.Range(FoundAt).Activate
            EndCellX = ActiveCell.Offset(-1, 0).Address
            EndCellY = ActiveCell.Offset(-1, 1).Address
            
            NumEntries = Sheet3.Range(BeginCellX, EndCellX).Rows.count
            Loop
                        
        If NumEntries >= 3 Then
            functStr = "=LINEST(" & BeginCellY & ":" & EndCellY & ", " & BeginCellX & ":" & EndCellX & "^{1,2}, TRUE, TRUE)" 'assumes a second order approximation to data
            vCoef = Evaluate(functStr) 'regresses parameters
            xmin = -0.5 * vCoef(1, 2) / vCoef(1, 1) 'evaluates the solution to {first derivative} = 0 = 2*vCoef(1,1)*xmin + vCoef(1,2)
            FreqStep = (Range(EndCellX).Value - Range(BeginCellX).Value) / (vCoef(4, 2) + 2) 'approximation to frequency step, where vCoef(4,2) = degress of freedom
            xminRound = 0 'initialize to default of zero in case the following gives an error
            On Error Resume Next
            xminRound = WorksheetFunction.MRound(-0.5 * vCoef(1, 2) / vCoef(1, 1), FreqStep) 'rounds the min frequency to the same precision as the frequency measurements
            
            Sheet5.Activate
            Range("A1").End(xlDown).Select 'bug over here that will go to the last possible row (A1048576) if there is only one row of text initially; hence adding units of the variables in the second row
            ActiveCell.Offset(1, 0).Value = Sheets(Sheets.count).Name 'the file name of the .csv file (which is currently the name of the data tab) is entered in cell e.g. A4
            ActiveCell.Offset(1, 1).Value = xminRound
            ActiveCell.Offset(1, 2).Value = FreqStep
            ActiveCell.Offset(1, 3).Value = vCoef(3, 1) 'R^2 of regression
            ActiveCell.Offset(1, 5).Value = vCoef(1, 1) 'x2 term coefficient
            ActiveCell.Offset(1, 6).Value = vCoef(1, 2) 'x term coefficient
            ActiveCell.Offset(1, 7).Value = vCoef(1, 3) 'y-intercept
            ActiveCell.Offset(1, 8).Value = vCoef(2, 1) 'stand. error of x2 term coefficient
            ActiveCell.Offset(1, 9).Value = vCoef(2, 2) 'stand. error of x term coefficient
            ActiveCell.Offset(1, 10).Value = vCoef(2, 3) 'stand. error of y-intercept
            ActiveCell.Offset(1, 11).Value = vCoef(3, 2) 'stand. error of y estimate
            ActiveCell.Offset(1, 12).Value = vCoef(4, 1) 'F of regression
            ActiveCell.Offset(1, 13).Value = vCoef(4, 2) 'degrees of freedom
            ActiveCell.Offset(1, 14).Value = vCoef(5, 1) 'SS of regression
            ActiveCell.Offset(1, 15).Value = vCoef(5, 2) 'SS of residuals
                      
            BeginCellX = FoundAt 'reintialize before searching for the next instance of Data
            End If
    End If

EndOfLoop:

    Loop While EndOfCol = "No"
        
End Sub
