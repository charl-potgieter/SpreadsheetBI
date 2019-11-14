Attribute VB_Name = "m030_FileUtilities"
Option Explicit
Option Private Module
'-----------------------------------------------------------------------------
'   Requires reference to Microsoft Scripting runtime
'-----------------------------------------------------------------------------


Function ReadTextFileIntoString(sFilePath As String) As String
'Inspired by:
'https://analystcave.com/vba-read-file-vba/

    Dim iFileNo As Integer
    
    'Get first free file number
    iFileNo = FreeFile

    Open sFilePath For Input As #iFileNo
    ReadTextFileIntoString = Input$(LOF(iFileNo), iFileNo)
    Close #iFileNo

End Function


Function WriteStringToTextFile(ByVal sStr As String, ByVal sFilePath As String)
'Requires reference to Microsoft Scripting Runtime
'Writes sStr to a text file

    Dim fso As Object
    Dim oFile As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(sFilePath)
    oFile.Write (sStr)
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing

End Function

Function FileNameFromPath(ByVal sFilePath As String) As String
    
    Dim fso As New FileSystemObject
    FileNameFromPath = fso.GetFileName(sFilePath)

End Function

Function FileNameFromPathExclExtension(ByVal sFilePath As String) As String
    
    Dim fso As New FileSystemObject
    FileNameFromPathExclExtension = fso.GetBaseName(sFilePath)

End Function




Sub FileItemsInFolder(ByVal sFolderPath As String, ByVal bRecursive As Boolean, ByRef FileItems() As Scripting.File)
'Requires refence: Microsoft Scripting Runtime
'Returns an array of files (which can be used to get filename, path etc)
'(Cannot create function due to recursive nature of the code)

    
    Dim fso As Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder
    Dim SubFolder As Scripting.Folder
    Dim FileItem As Scripting.File
    
    Set fso = New Scripting.FileSystemObject
    Set SourceFolder = fso.GetFolder(sFolderPath)
    
    For Each FileItem In SourceFolder.Files
    
        If Not ArrayIsInitialised(FileItems) Then
            ReDim FileItems(0)
        Else
            ReDim Preserve FileItems(UBound(FileItems) + 1)
        End If
        
        Set FileItems(UBound(FileItems)) = FileItem
        
    Next FileItem
    
    If bRecursive Then
        For Each SubFolder In SourceFolder.SubFolders
            FileItemsInFolder SubFolder.Path, True, FileItems
        Next SubFolder
    End If
    
    Set FileItem = Nothing
    Set SourceFolder = Nothing
    Set fso = Nothing
    

End Sub



Function FileIsOpen(ByVal sFilePath) As Boolean
'Requires refence: Microsoft Scripting Runtime

    Dim fso As New FileSystemObject
    Dim sFileName As String
    Dim wkb As Workbook
    
    sFileName = fso.GetFileName(sFilePath)
    
    On Error Resume Next
    Set wkb = Workbooks(sFileName)
    FileIsOpen = (Err.Number = 0)
    On Error GoTo 0

End Function


Function NumberOfFilesInFolder(ByVal sFolderPath As String) As Integer
'Requires refence: Microsoft Scripting Runtime
'This is non-recursive


    Dim oFSO As FileSystemObject
    Dim oFolder As Folder
    
    Set oFSO = New FileSystemObject
    Set oFolder = oFSO.GetFolder(sFolderPath)
    NumberOfFilesInFolder = oFolder.Files.Count


End Function


Function GetFolder() As String
'Returns the results of a user folder picker

    Dim fldr As FileDialog
    Dim sItem As String
    
    Set fldr = Application.FileDialog(msoFileDialogFilePicker)
    With fldr
        .Title = "Select a folder"
        .AllowMultiSelect = False
        .InitialFileName = ThisWorkbook.Path
        If .Show = -1 Then GoTo NextCode
        sItem = .SelectedItems(1)
    End With
    
NextCode:
    GetFolder = sItem
    Set fldr = Nothing


End Function




Function GetSelectedfiles(Optional sDefaultFolder As String) As String()
'Returns selected file paths from users as an array of strings
    
    Dim ArrSelectedValues() As String
    Dim i As Integer
    
    With Application.FileDialog(msoFileDialogOpen)
        .AllowMultiSelect = True
        .Show
        If IsMissing(sDefaultFolder) Then
            .InitialFileName = sDefaultFolder
        End If
        
        ReDim ArrSelectedValues(.SelectedItems.Count - 1)
        For i = 0 To (.SelectedItems.Count - 1)
            ArrSelectedValues(i) = .SelectedItems(i + 1)
        Next i
    End With
        
    GetSelectedfiles = ArrSelectedValues

End Function























