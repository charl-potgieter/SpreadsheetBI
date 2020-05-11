Attribute VB_Name = "m030_FileUtilities"
Option Explicit
Option Private Module
'-----------------------------------------------------------------------------
'   Requires reference to Microsoft Scripting runtime
'-----------------------------------------------------------------------------

Function FolderExists(ByVal sFolderPath) As Boolean
'Requires reference to Microsoft Scripting runtime
'An alternative solution exists using the DIR function but this seems to result in memory leak and folder is
'not released by VBA
    
    Dim FSO As Scripting.FileSystemObject
    Dim FolderPath As String
    
    Set FSO = New Scripting.FileSystemObject
    
    If Right(sFolderPath, 1) <> Application.PathSeparator Then
        FolderPath = FolderPath & Application.PathSeparator
    End If
    
    FolderExists = FSO.FolderExists(sFolderPath)
    Set FSO = Nothing

End Function

Function FileExists(ByVal sFilePath) As Boolean
'Requires reference to Microsoft Scripting runtime
'An alternative solution exists using the DIR function but this seems to result in memory leak and file is
'not released by VBA

    Dim FSO As Scripting.FileSystemObject
    Dim FolderPath As String
    
    Set FSO = New Scripting.FileSystemObject
    
    FileExists = FSO.FileExists(sFilePath)
    Set FSO = Nothing


End Function


Sub CreateFolder(ByVal sFolderPath As String)
'   Requires reference to Microsoft Scripting runtime

    Dim FSO As FileSystemObject

    If FolderExists(sFolderPath) Then
        MsgBox ("Folder already exists, new folder not created")
    Else
        Set FSO = New FileSystemObject
        FSO.CreateFolder sFolderPath
    End If
    
    Set FSO = Nothing

End Sub



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

    Dim FSO As Object
    Dim oFile As Object
    
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set oFile = FSO.CreateTextFile(sFilePath)
    oFile.Write (sStr)
    oFile.Close
    Set FSO = Nothing
    Set oFile = Nothing

End Function

Function FileNameFromPath(ByVal sFilePath As String) As String
    
    Dim FSO As New FileSystemObject
    FileNameFromPath = FSO.GetFileName(sFilePath)

End Function

Function FileNameFromPathExclExtension(ByVal sFilePath As String) As String
    
    Dim FSO As New FileSystemObject
    FileNameFromPathExclExtension = FSO.GetBaseName(sFilePath)

End Function




Sub FileItemsInFolder(ByVal sFolderPath As String, ByVal bRecursive As Boolean, ByRef FileItems() As Scripting.File)
'Returns an array of files (which can be used to get filename, path etc)
'(Cannot create function due to recursive nature of the code)

    
    Dim FSO As Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder
    Dim SubFolder As Scripting.Folder
    Dim FileItem As Scripting.File
    
    Set FSO = New Scripting.FileSystemObject
    Set SourceFolder = FSO.GetFolder(sFolderPath)
    
    For Each FileItem In SourceFolder.Files
    
        If Not ArrayIsDimensioned(FileItems) Then
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
    Set FSO = Nothing
    

End Sub



Function FileIsOpen(ByVal sFilePath) As Boolean
'Requires refence: Microsoft Scripting Runtime

    Dim FSO As New FileSystemObject
    Dim sFileName As String
    Dim wkb As Workbook
    
    sFileName = FSO.GetFileName(sFilePath)
    
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
    
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    With fldr
        .Title = "Select a folder"
        .AllowMultiSelect = False
        .InitialFileName = ActiveWorkbook.Path
        If .Show = -1 Then
            GetFolder = .SelectedItems(1)
        End If
    End With
    
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




Sub ExportWorksheetSheetToPipeDelimtedText(ByRef sht As Worksheet, ByVal sFilePathAndName As String)
'Requires reference to Microsoft Scripting Runtime
'Saves sht as a pipe delimted text file
'Existing files will not be overwritten.  Warning is given.
    
    Dim dblNumberOfRows As Double
    Dim dblNumberOfCols As Double
    Dim iFileNo As Integer
    Dim i As Double
    Dim j As Double
    Dim sRowStringToWrite As String

    
    If FileExists(sFilePathAndName) Then
        MsgBox ("File " & sFilePathAndName & " already exists.  New file has not been generated")
    End If

    'Get first free file number
    iFileNo = FreeFile

    
    Open sFilePathAndName For Output As #iFileNo
    
    dblNumberOfRows = ActiveSheet.UsedRange.Rows.Count
    dblNumberOfCols = ActiveSheet.UsedRange.Columns.Count
    
    
    For j = 1 To dblNumberOfRows
        sRowStringToWrite = ""
        For i = 1 To dblNumberOfCols
            If i < dblNumberOfCols Then
                sRowStringToWrite = sRowStringToWrite & sht.Cells(j, i) & "|"
            Else
                sRowStringToWrite = sRowStringToWrite & sht.Cells(j, i)
            End If
        Next i
        Print #1, sRowStringToWrite
    Next j

    Close #1

End Sub




Sub ExportListObjectToPipeDelimtedText(ByRef lo As ListObject, ByVal sFilePathAndName As String)
'Requires reference to Microsoft Scripting Runtime
'Saves sht as a pipe delimted text file
'Existing files will not be overwritten.  Warning is given.
    
    Dim dblNumberOfRows As Double
    Dim dblNumberOfCols As Double
    Dim iFileNo As Integer
    Dim i As Double
    Dim j As Double
    Dim sRowStringToWrite As String

    
    If FileExists(sFilePathAndName) Then
        MsgBox ("File " & sFilePathAndName & " already exists.  New file has not been generated")
    End If

    'Get first free file number
    iFileNo = FreeFile

    
    Open sFilePathAndName For Output As #iFileNo
    
    dblNumberOfRows = lo.Range.Rows.Count
    dblNumberOfCols = lo.Range.Columns.Count
    
    
    For j = 1 To dblNumberOfRows
        sRowStringToWrite = ""
        For i = 1 To dblNumberOfCols
            If i < dblNumberOfCols Then
                sRowStringToWrite = sRowStringToWrite & lo.Range.Cells(j, i) & "|"
            Else
                sRowStringToWrite = sRowStringToWrite & lo.Range.Cells(j, i)
            End If
        Next i
        If j < dblNumberOfRows Then
            Print #iFileNo, sRowStringToWrite
        Else
            'note the semi-colon at end to avoid the newline
            Print #iFileNo, sRowStringToWrite;
        End If
    Next j

    Close #iFileNo

End Sub




Function GetNextAvailableFileName(ByVal sFilePath As String) As String
'Requires refence: Microsoft Scripting Runtime
'Returns next available file name.  Can be utilised to ensure files are not overwritten

    Dim oFSO As FileSystemObject
    Dim sFolder As String
    Dim sFileName As String
    Dim sFileExtension As String
    Dim i As Long

    Set oFSO = CreateObject("Scripting.FileSystemObject")

    With oFSO
        sFolder = .GetParentFolderName(sFilePath)
        sFileName = .GetBaseName(sFilePath)
        sFileExtension = .GetExtensionName(sFilePath)

        Do While .FileExists(sFilePath)
            i = i + 1
            sFilePath = .BuildPath(sFolder, sFileName & "(" & i & ")." & sFileExtension)
        Loop
        
    End With

    GetNextAvailableFileName = sFilePath

End Function

