Attribute VB_Name = "LoopThroughFolder"

'Option Explicit

Sub LoopThroughFold()

'compiled August 2020

'This sub works through the .csv data files in the folder specified (in "DataFilePath"), where the recorded data is voltage vs. frequency profiles. The frequency at the minimum voltage is determined by approximating the profile as a quadratic curve, and the frequency as well as some stats are returned.
'The .csv files have the file name format: "AAAAAA_BB_CCCC.csv", where
'    AAAAAA - code related to the chemical system being investigated
'    BB - set point temperature for measured data, in degrees C
'    CCCC - composition of chemical system being investigated, as a mass fraction
'        e.g. C12OH1_35_0055.csv is for data measured in the binary 1-dodecanol & carbon dioxide system at 35 C, with an alcohol mass fraction of 0.0055

'Data in the .csv file is the frequency and voltage, with a note at the beginning of each data block indicating the date and time at which data collection started: "Data collection started at: Thu Jul 04 14:31:54 2019"
'A space is used as delimiter between the frequency and voltage
'Multiple periods of data collection may be in one file, but all data will be for the same chemical system, composition & temperature

'Frequency min.xlsm is the main file where the output will be recorded; it contains the following three tabs:
'   Sheet1 ("Home") - starting tab with buttons for code execution: a button executes sub "GetFolderPathNFileCount" (output as folder path and number of files in cells F8:F9 on "Home" tab), and another button executes sub "LoopThroughFold" (output as stats on "Results" tab)
'   Sheet3 ("Worksheet") - working space used to copy data from the .csv files
'   Sheet5 ("Results") - records stats from the quadratic regression and estimate of frequency; has the following columns: Filename (A), Min Freq [Hz] (B), FreqStep [Hz] (C), R^2 (D), A (x2 coefficient) (F), B (x coefficient) (G), C (y intercept) (H), Standard error of A (I), Standard error of B (J), Standard error of C (K), Standard error of Y estimate (L), F statistic (M), degrees of freedom (N), Sum of squares of the regression (O), and Sum of squares of the residuals (P)

Dim fldr As FileDialog
Dim sItem As String
Dim x As Integer

Application.ScreenUpdating = False 'stops screen from flashing while executing code
DataFilePath = Sheet1.Range("F8").Value
Datafile = Dir(DataFilePath)
x = 0 'counts the number of files successfully copied

If Datafile = "" Then
    MsgBox "Please choose folder path first!", vbCritical
    Exit Sub
    Else
        Sheet5.Range("A1").End(xlDown).Offset(1, 0).Value = "Folder: " & DataFilePath 'records the folder path for the files that are about to be processed on the "Results" tab
    End If

Do While Len(Datafile) > 0
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Workbooks.Open (DataFilePath & Datafile)
    Sheets(1).Copy after:=Workbooks("Frequency min.xlsm").Sheets(Workbooks("Frequency min.xlsm").Sheets.count) 'copies the contents of the selected .csv file onto a new tab after the "Results" tab
    
    Workbooks(Datafile).Close
    x = x + 1
    Sheets(Sheets.count).Activate 'assuming this workbook is the only Excel file open
    Filename = Datafile
    Sheets(Sheets.count).Name = Filename 'renames the new tab containing data from the .csv file
    
    Sheets(Sheets.count).Activate
    Range("A1", Range("A1").End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheet3.Activate
    Range("A1").Select
    ActiveSheet.Paste 'paste the data from the newly created tab onto the "Worksheet" tab; copying directly from the .csv file previously crashed the program
    
    RowCount = Sheet3.Range("A1", Range("A1").End(xlDown)).Rows.count 'verifies that there is recorded data on the .csv file under consideration
    rCounter = 1
    Do While rCounter <= RowCount 'this loop will indicate individual data blocks/periods of data collection by resulting in either "Data" or the frequency in column G, and voltage or "collection started at..." in column H
        Sheet3.Cells(rCounter, 6).Value = "=IFERROR(FIND("" "",A" & rCounter & "),5)"
        Sheet3.Cells(rCounter, 7).Value = "=IFERROR(VALUE(LEFT(A" & rCounter & ",F" & rCounter & ")),LEFT(A" & rCounter & ",F" & rCounter & "))"
        Sheet3.Cells(rCounter, 8).Value = "=IFERROR(VALUE(RIGHT(A" & rCounter & ",LEN(A" & rCounter & ")-F" & rCounter & ")), RIGHT(A" & rCounter & ",LEN(A" & rCounter & ")-F" & rCounter & "))"
        rCounter = rCounter + 1
        Loop
    
    Range("G1", Range("H1").End(xlDown)).Select
    Application.CutCopyMode = False
    Selection.Copy
    Range("L1").PasteSpecial xlPasteValues 'the values in columns G & H are copied to columns L & M for further processing
    
    Call FindMinFreq 'the Sub "FindMinFreq" is called at this point to determine the frequency and corresponding stats
    
    Sheets(Sheets.count).Delete 'deletes the tab with .csv data and clears the "Worksheet" tab
    Sheet3.Activate
    Sheet3.Cells.Select
    Selection.Value = ""
    
    Datafile = Dir 'when called again with no further arguments, just returns the next file in the folder/directory
    Sheet1.Activate
    Application.ScreenUpdating = True
Loop

Sheet1.Activate
Application.ScreenUpdating = True

MsgBox "Done with analysis", vbInformation
'Sheet1.Range("B8:F9").Value = ""

End Sub
