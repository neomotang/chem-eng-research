Attribute VB_Name = "f4Analyse1Day"

Sub Analyse1Day()

'compiled August 2024

'This sub processes data from a single day, determining the average values at specified times from the .csv files. These 3 .csv files with time stamped data are of interest:
'   PrTemp.csv - data recorded every 10 seconds, with columns A - I containing: A - time, B - date, columns C-I - temperature measurements T1-1, T2-1, T3-1, T3-2, T3-3, T4-1 & T4-2, respectively
'   PrFlow.csv - data recorded every 10 seconds, with columns A - F containing: A - time, B - date, C - differential pressure, D - flow rate, columns E&F - pressure measurements P4-1 & P4-2, respectively
'   PrDp.csv - data recorded approximately every 1/5 of a second, with columns A - C containing: A - time, B - data, C - differential pressure

'AnalyseCSV.xlsm is the main spreadsheet where data are processed and saved; it contains the following three tabs:
'   "Home" (Sheet1) - starting tab with prompts, several buttons for executing code, and outputs from code execution
'       Row 21 onwards has user input data for each experimental run (date, start and end time in columns A-C) to be analysed, and results of the analysis are saved in column D onwards
'   "Worksheet" (Sheet2) - workspace for copying data and calculations
'   "Results" (Sheet3) - processed data are saved here, copying results from the "Home" tab

Dim fldr As FileDialog
Dim sItem As String
Dim x As Integer

Application.ScreenUpdating = False
Application.DisplayAlerts = False

DataCount = Sheet1.Range("A20", Range("A20").End(xlDown)).Rows.Count 'count how many experimental runs are still left to be analysed

If DataCount <= 1 Or DataCount = 1048557 Then '1 = no entries, i.e. just heading row contains text; 1048557 = no entries but Excel has also selected all the blank rows til bottom of spreadsheet
    Exit Sub 'leave without doing anything; also no warning message displayed
Else
    Tempfile = "PrTemp.csv"
    Flowfile = "PrFlow.csv"
    DPfile = "PrDp.csv"
    
    fCounter = 21 'row counter that represents experimental run entries, starting with the first entry in row 21
    DataFilePath = Sheet1.Range("F8").Value & "CSV\" 'the .csv files of interest are in the "CSV" subfolder, with the day's folder stored in cell F8

    'Open the PrTemp.csv file, paste it as a new tab at the end of the spreadsheet, and then paste the contents to start in column C of the "Worksheet" tab
    Workbooks.Open (DataFilePath & Tempfile)
    Sheets(1).Copy after:=Workbooks("AnalyseCSV.xlsm").Sheets(Workbooks("AnalyseCSV.xlsm").Sheets.Count)
    Workbooks(Tempfile).Close
    Sheets(Sheets.Count).Activate 'assuming this workbook is the only Excel file open
    Range("A1", Range("I1").End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheet2.Activate
    Range("C1").Select
    ActiveSheet.Paste
       
    'Open the PrFlow.csv file, paste it as a new tab at the end of the spreadsheet, and then paste the contents to start in column L of the "Worksheet" tab
    Workbooks.Open (DataFilePath & Flowfile)
    Sheets(1).Copy after:=Workbooks("AnalyseCSV.xlsm").Sheets(Workbooks("AnalyseCSV.xlsm").Sheets.Count)
    Workbooks(Flowfile).Close
    Sheets(Sheets.Count).Activate 'assuming this workbook is the only Excel file open
    Range("A1", Range("F1").End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheet2.Activate
    Range("L1").Select
    ActiveSheet.Paste
    
    'Open the PrDp.csv file, paste it as a new tab at the end of the spreadsheet, and then paste the contents to start in column AC of the tab "Worksheet"
    Workbooks.Open (DataFilePath & DPfile)
    Sheets(1).Copy after:=Workbooks("AnalyseCSV.xlsm").Sheets(Workbooks("AnalyseCSV.xlsm").Sheets.Count)
    Workbooks(DPfile).Close
    Sheets(Sheets.Count).Activate 'assuming this workbook is the only Excel file open
    Range("A1", Range("C1").End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheet2.Activate
    Range("AC1").Select
    ActiveSheet.Paste

    RowCount = Sheet2.Range("C1", Range("Q1").End(xlDown)).Rows.Count 'counter representing the total data rows in the PrTemp.csv and PrFlow.csv files
    
    RCounter = 2 'counter used to work through the recorded data rows in PrTemp.csv and PrFlow.csv; row 1 is the column headings

    Do While RCounter <= RowCount 'DoLoop#1 - Determine the approximate start time in the .csv files from the difference between entries in the time column (C) and the user input experimental start time on the "Home" tab (Sheet1, column B)
        
        Sheet2.Cells(RCounter, 1).Value = "=C" & RCounter & "-Home!$B$" & fCounter 'calculation in column A
        timeDif = Sheet2.Cells(RCounter, 1).Value
        
        If timeDif < 0 Then 'the difference in time will be negative until the user input experimental start time is reached
            RCounter = RCounter + 1
        Else
            timeDifPrev = Sheet2.Cells(RCounter - 1, 1).Value
            If (timeDif + timeDifPrev) < 0 Then 'choose the approximate time as whichever is closest to the recorded start time as on the "Home" tab
                approxStart = Sheet2.Cells(RCounter, 3).Value
            Else
                approxStart = Sheet2.Cells(RCounter - 1, 3).Value
                RCounter = RCounter - 1 'if the closest time is in the previous row, move the row counter back up a single row, in preparation for calculating averages in the next lines of code
            End If
            
            'Print the approximate start time to column D on the "Home" (Sheet1) tab with appropriate formatting
            Sheet1.Cells(fCounter, 4) = approxStart
            Sheet1.Cells(fCounter, 4).NumberFormat = "hh:mm:ss"
            
            'Use the row number for the approximate start time and get average values for data in PrTemp.csv & PrFlow.csv around that time; later print to "Home" tab with appropriate formatting
            Call GetAvgTemps(RCounter) 'calls routine in "f50CalcAvgT.bas" which calculates averages for PrTemp.csv and stores them in public variables
            
            Call PrintStartTemps(fCounter) 'calls routine in "f60PrintAvgStartT" which saves the values in the public variables to the "Home" tab
            
            Call GetAvgFlows(RCounter) 'calls routine in "f51CalcAvgF.bas" which calculates averages for PrFlow.csv and stores them in public variables
            
            Call PrintStartFlows(fCounter) 'calls routine in "f61PrintAvgStartF.bas" which saves the values in the public variables to the "Home" tab
            
            Do While RCounter <= RowCount 'DoLoop#2 - Determine the approximate end time in the .csv files from the difference between entries in the time column (C) and the user input experimental end time on the "Home" tab (Sheet1, column C)
                
                Sheet2.Cells(RCounter, 2).Value = "=C" & RCounter & "-Home!$C$" & fCounter 'calculation in column B
                timeDifE = Sheet2.Cells(RCounter, 2).Value
                
                If timeDifE < 0 Then 'the difference in time will be negative until the user input experimental end time is reached
                    RCounter = RCounter + 1
                Else
                    timeDifPrevE = Sheet2.Cells(RCounter - 1, 2).Value
                    If (timeDifE + timeDifPrevE) < 0 Then 'choose the approximate time as whichever is closest to the recorded end time as on the "Home" tab
                        approxEnd = Sheet2.Cells(RCounter, 3).Value
                    Else
                        approxEnd = Sheet2.Cells(RCounter - 1, 3).Value
                        RCounter = RCounter - 1 'if the closest time is in the previous row, move the row counter back up a single row, in preparation for calculating averages in the next lines of code
                    End If
                    
                    'Print the approximate end time to column T on the "Home" (Sheet1) tab with appropriate formatting
                    Sheet1.Cells(fCounter, 20).Value = approxEnd
                    Sheet1.Cells(fCounter, 20).NumberFormat = "hh:mm:ss"
                    
                     'Use the row number for the approximate end time and get average values for data in PrTemp.csv & PrFlow.csv around that time; later print to "Home" tab with appropriate formatting
                    Call GetAvgTemps(RCounter) 'calls routine in "f50CalcAvgT.bas" which calculates averages for PrTemp.csv and stores them in public variables
                    
                    Call PrintEndTemps(fCounter) 'calls routine in "f70PrintAvgEndT" which saves the values in the public variables to the "Home" tab
                    
                    Call GetAvgFlows(RCounter) 'calls routine in "f51CalcAvgF.bas" which calculates averages for PrFlow.csv and stores them in public variables
                    
                    Call PrintEndFlows(fCounter) 'calls routine in "f71PrintAvgEndF.bas" which saves the values in the public variables to the "Home" tab
                                            
                    RCounter = RCounter + 1
                    Exit Do 'exits DoLoop#2
                End If
                Loop
            Exit Do 'exits DoLoop#1
        End If
    Loop
        
    dpRowCount = Sheet2.Range("AC1", Range("AE1").End(xlDown)).Rows.Count 'counter representing the total data rows in the PrDp.csv file, since its time intervals are different
        
    RCounter = 2 'counter used to work through the recorded data rows in PrDp.csv, does not need to be unique; row 1 is the column headings
    
    Do While RCounter <= dpRowCount 'DoLoop#3 - Determine the approximate start time in the .csv file from the difference between entries in the time column (AC) and the user input experimental start time on the "Home" tab (Sheet1, column B)
    
        Sheet2.Cells(RCounter, 27).Value = "=AC" & RCounter & "-Home!$B$" & fCounter 'calculation in column AA
        timeDif = Sheet2.Cells(RCounter, 27).Value
        
        If timeDif < 0 Then 'the difference in time will be negative until the user input experimental start time is reached
            RCounter = RCounter + 1
        Else
            timeDifPrev = Sheet2.Cells(RCounter - 1, 27).Value
            If (timeDif + timeDifPrev) < 0 Then 'choose the approximate time as whichever is closest to the recorded start time as on the "Home" tab
                approxStart = Sheet2.Cells(RCounter, 29).Value
            Else
                approxStart = Sheet2.Cells(RCounter - 1, 29).Value
                RCounter = RCounter - 1 'if the closest time is in the previous row, move the row counter back up a single row, in preparation for calculating averages in the next lines of code
            End If
            
            'Print the approximate start time to column Q on the "Home" (Sheet1) tab with appropriate formatting
            Sheet1.Cells(fCounter, 17) = approxStart
            Sheet1.Cells(fCounter, 17).NumberFormat = "hh:mm:ss"
            
            'Use the row number for the approximate start time and get average values for data in PrDp.csv around that time; later print to "Home" tab with appropriate formatting
            Call GetAvgDP(RCounter) 'calls routine in "f52CalcAvgDP.bas" which calculates averages for PrDp.csv and stores them in public variables
            
            Call PrintStartDP(fCounter) 'calls routine in "f62PrintAvgStartDP" which saves the values in the public variables to the "Home" tab
            
            Do While RCounter <= dpRowCount 'DoLoop#4 - Determine the approximate end time in the .csv file from the difference between entries in the time column (AC) and the user input experimental end time on the "Home" tab (Sheet1, column C)
                
                Sheet2.Cells(RCounter, 28).Value = "=AC" & RCounter & "-Home!$C$" & fCounter 'calculation in column AB
                timeDifE = Sheet2.Cells(RCounter, 28).Value
                
                If timeDifE < 0 Then 'the difference in time will be negative until the user input experimental end time is reached
                    RCounter = RCounter + 1
                Else
                    timeDifPrevE = Sheet2.Cells(RCounter - 1, 28).Value
                    If (timeDifE + timeDifPrevE) < 0 Then 'choose the approximate time as whichever is closest to the recorded end time as on the "Home" tab
                        approxEnd = Sheet2.Cells(RCounter, 29).Value
                    Else
                        approxEnd = Sheet2.Cells(RCounter - 1, 29).Value
                        RCounter = RCounter - 1 'if the closest time is in the previous row, move the row counter back up a single row, in preparation for calculating averages in the next lines of code
                    End If
                    
                    'Print the approximate end time to column AG on the "Home" (Sheet1) tab with appropriate formatting
                    Sheet1.Cells(fCounter, 33).Value = approxEnd
                    Sheet1.Cells(fCounter, 33).NumberFormat = "hh:mm:ss"
                    
                    'Use the row number for the approximate end time and get average values for data in PrDp.csv around that time; later print to "Home" tab with appropriate formatting
                    Call GetAvgDP(RCounter) 'calls routine in "f52CalcAvgDP.bas" which calculates averages for PrDp.csv and stores them in public variables
                    
                    Call PrintEndDP(fCounter) 'calls routine in "f72PrintAvgEndDP" which saves the values in the public variables to the "Home" tab
                    
                    RCounter = RCounter + 1
                    Exit Do 'exits DoLoop#4
                End If
            
            Loop
            Exit Do 'exits DoLoop#3
        End If
    
    Loop
        
End If

Sheets(Sheets.Count).Delete 'delete "PrDp" tab
Sheets(Sheets.Count).Delete 'delete "PfFlow" tab
Sheets(Sheets.Count).Delete 'delete "PrTemp" tab
Sheet2.Activate
Sheet2.Cells.Select
Selection.Value = "" 'clear "Worksheet" tab
Sheet1.Activate
Application.ScreenUpdating = True
Application.DisplayAlerts = True

MsgBox "Done with analysis", vbInformation

End Sub

