Attribute VB_Name = "f1GetGroupFolder"

Sub GetFolderPathNFileCount()

'compiled August 2024

'This sub gets the folder path containing multiple days' data in individual subfolders, and also counts the number of .csv files in the folders to be analysed
'The folder path, number of subfolders and number of files are saved to cells F8 - F10 on the "Home" tab (Sheet1)
    
Dim fldr As FileDialog

Set fldr = Application.FileDialog(msoFileDialogFolderPicker) 'choose folder containing subfolders with individual days' data
fldr.Title = "Select the folder containing your .csv data"
fldr.AllowMultiSelect = False
fldr.InitialFileName = strPath

If fldr.Show <> -1 Then
    MsgBox "Folder was not chosen. Please try again.", vbCritical
    Else:
        FolderPath = fldr.SelectedItems(1) 'retrieves folder path name
        folderCount = CountFolders(fldr.SelectedItems(1)) 'calls private function "CountFolders" that counts the number of subfolders
        FileCount = CountFiles(fldr.SelectedItems(1)) 'calls private function "CountFiles" that counts the number of files
        Sheet1.Activate 'Return to the "Home" tab to save the details
        Sheet1.Range("B8").Value = "Folder path chosen:"
        Sheet1.Range("F8").Value = FolderPath & "\"
        Sheet1.Range("B9").Value = "Number of folders in folder:"
        Sheet1.Range("F9").Value = folderCount
        Sheet1.Range("B10").Value = "Number of files in folder:"
        Sheet1.Range("F10").Value = FileCount
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
    Set objFiles = objFso.GetFolder(strDirectory).Files

    'Count files (that match the extension if provided)
    If strExt = "*.*" Then
        CountFiles = objFiles.Count
    Else
        For Each objFile In objFiles
            If UCase(Right(objFile.Path, (Len(objFile.Path) - InStrRev(objFile.Path, ".")))) = UCase(strExt) Then
                CountFiles = CountFiles + 1
            End If
        Next objFile
    End If

EarlyExit:
    On Error Resume Next
    Set objFile = Nothing
    Set objFiles = Nothing
    Set objFso = Nothing
    On Error GoTo 0
End Function
Private Function CountFolders(directoryPath As String) As Long
    Dim folder As Object
    Dim folderCount As Long
    
    ' Create a FileSystemObject
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Get the specified folder
    Dim parentFolder As Object
    Set parentFolder = fso.GetFolder(directoryPath)
    
    ' Loop through each subfolder and increment the count
    For Each folder In parentFolder.Subfolders
        folderCount = folderCount + 1
    Next folder
    
    ' Return the total folder count
    CountFolders = folderCount
End Function
