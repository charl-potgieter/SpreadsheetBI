Attribute VB_Name = "m030_FileUtilities"
Option Explicit

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
