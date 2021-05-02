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
    
    Dim fso As Scripting.FileSystemObject
    Dim FolderPath As String
    
    Set fso = New Scripting.FileSystemObject
    
    If Right(sFolderPath, 1) <> Application.PathSeparator Then
        FolderPath = FolderPath & Application.PathSeparator
    End If
    
    FolderExists = fso.FolderExists(sFolderPath)
    Set fso = Nothing

End Function

Function FileExists(ByVal sFilePath) As Boolean
'Requires reference to Microsoft Scripting runtime
'An alternative solution exists using the DIR function but this seems to result in memory leak and file is
'not released by VBA

    Dim fso As Scripting.FileSystemObject
    Dim FolderPath As String
    
    Set fso = New Scripting.FileSystemObject
    
    FileExists = fso.FileExists(sFilePath)
    Set fso = Nothing


End Function


Sub CreateFolder(ByVal sFolderPath As String)
'   Requires reference to Microsoft Scripting runtime

    Dim fso As FileSystemObject

    If FolderExists(sFolderPath) Then
        MsgBox ("Folder already exists, new folder not created")
    Else
        Set fso = New FileSystemObject
        fso.CreateFolder sFolderPath
    End If
    
    Set fso = Nothing

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
'*** THIS WILL OVERWRITE ANY CURRENT CONTENT OF THE FILE ***

    Dim fso As Object
    Dim oFile As Object
    
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(sFilePath)
    oFile.Write (sStr)
    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing

End Function


Function WriteStringToPipeDelimitedTextFileAddQuotes(ByVal sStr As String, ByVal sFilePath As String)
'Requires reference to Microsoft Scripting Runtime
'Writes sStr to a text file where sStr is a pipe delimited string
'This sub adds double quotes to each delimited item.  This is useful when there are newlines in the string.
'The resulting file can then be picked up by power query by setting QuoteStyle.Csv

    Dim fso As Object
    Dim oFile As Object
    
    'Add double quotes to start and end of string
    sStr = """" & sStr & """"
    
    'Add double quotes to each pipe delimitter
    sStr = Replace(sStr, "|", """|""")
    
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
'Returns an array of files (which can be used to get filename, path etc)
'(Cannot create function due to recursive nature of the code)

    
    Dim fso As Scripting.FileSystemObject
    Dim SourceFolder As Scripting.Folder
    Dim SubFolder As Scripting.Folder
    Dim FileItem As Scripting.File
    
    Set fso = New Scripting.FileSystemObject
    Set SourceFolder = fso.GetFolder(sFolderPath)
    
    For Each FileItem In SourceFolder.Files
    
        If Not IsArrayAllocated(FileItems) Then
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




Sub DeleteFirstLineOfTextFile(ByVal sFilePathAndName As String)

    Dim sInput As String
    Dim sOutput() As String
    Dim i As Long
    Dim j As Long
    Dim lSizeOfOutput As Long
    Dim iFileNo As Integer
    
    iFileNo = FreeFile
    
    'Import file lines to array excluding firt line
    Open sFilePathAndName For Input As iFileNo
    i = 0
    j = 0
    Do Until EOF(iFileNo)
        j = j + 1
        Line Input #iFileNo, sInput
        If j > 1 Then
            i = i + 1
            ReDim Preserve sOutput(1 To i)
            sOutput(i) = sInput
        End If
    Loop
    Close #iFileNo
    lSizeOfOutput = i
    
    'Write array to file
    Open sFilePathAndName For Output As 1
    For i = 1 To lSizeOfOutput
        Print #iFileNo, sOutput(i)
    Next i
    Close #iFileNo
    
    
End Sub

