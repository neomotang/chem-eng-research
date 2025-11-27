Attribute VB_Name = "GetFolderPathNFileNum"

'Option Explicit

Sub GetFolderPathNFileCount()

'compiled August 2020

'To be used with LoopThroughFolder.bas and MinFrequency.bas
    
Dim fldr As FileDialog

Set fldr = Application.FileDialog(msoFileDialogFolderPicker) 'choose folder containing csv data
fldr.Title = "Select the folder containing your .csv data"
fldr.AllowMultiSelect = False
fldr.InitialFileName = strPath

If fldr.Show <> -1 Then
    MsgBox "Folder was not chosen. Please try again.", vbCritical 
    Exit Sub
    Else:
        FolderPath = fldr.SelectedItems(1) 'retrieves path name
        FileCount = CountFiles(fldr.SelectedItems(1))
        Debug.Print FileCount
        Sheet1.Activate
        MsgBox FileCount & " files found in folder " & FolderPath
        Sheet1.Range("B8").Value = "Folder path chosen:"
        Sheet1.Range("F8").Value = FolderPath & "\"
        Sheet1.Range("B9").Value = "Number of files in folder:"
        Sheet1.Range("F9").Value = FileCount
    End If

Sheet1.Activate
End Sub

Private Function CountFiles(strDirectory As String, Optional strExt As String = "*.*") As Double
'Function purpose: To count files in a directory.  If a file extension is provided,
'   then count only files of that type, otherwise return a count of all files.
    Dim objFso As Object
    Dim objFiles As Object
    Dim objFile As Object

    'Set Error Handling
    On Error GoTo EarlyExit

    'Create objects to get a count of files in the directory
    Set objFso = CreateObject("Scripting.FileSystemObject")
    Set objFiles = objFso.Getfolder(strDirectory).Files

    'Count files (that match the extension if provided)
    If strExt = "*.*" Then
        CountFiles = objFiles.count
    Else
        For Each objFile In objFiles
            If UCase(Right(objFile.Path, (Len(objFile.Path) - InStrRev(objFile.Path, ".")))) = UCase(strExt) Then
                CountFiles = CountFiles + 1
            End If
        Next objFile
    End If

EarlyExit:
    'Clean up
    On Error Resume Next
    Set objFile = Nothing
    Set objFiles = Nothing
    Set objFso = Nothing
    On Error GoTo 0
End Function
