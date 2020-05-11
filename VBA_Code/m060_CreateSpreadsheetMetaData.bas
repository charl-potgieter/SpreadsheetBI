Attribute VB_Name = "m060_CreateSpreadsheetMetaData"
Option Explicit
Option Private Module

Sub GenerateWorksheetMetaDataFile(ByRef wkb As Workbook)

'Generates selected spreadsheet data to allo the spreadsheet to be recreated
'via VBA.
'Aspects covered:
'   - Sheet names
'   - Sheet category
'   - Sheet heading
'   - Table name
'   - Number of table columns


    Dim sht As Worksheet
    Dim shtTemp As Worksheet
    Dim lo As ListObject
    Dim sSheetHeader As String
    Dim sSheetCategory As String
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer
    Dim sFilePathAndName As String

    Set wkb = ActiveWorkbook
    sFolderPath = wkb.Path & Application.PathSeparator & "SpreadsheetMetaData"
    sFilePathAndName = sFolderPath & Application.PathSeparator & "SpreadsheetMetaData.txt"

    If Not (FolderExists(sFolderPath)) Then
        CreateFolder (sFolderPath)
    End If
    
    sFilePathAndName = GetNextAvailableFileName(sFilePathAndName)

    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "Name|Sheet Category|Sheet Header|Table Name|Number Of Table Columns";
    
    For Each sht In wkb.Worksheets
        
        If sht.Name <> "Index" Then
        
            On Error Resume Next
            sSheetHeader = sht.Names("SheetHeading").RefersToRange.Value
            sSheetCategory = sht.Names("SheetCategory").RefersToRange.Value
            If Err.Number = 0 Then
                sRowToWrite = vbCr & _
                    sht.Name & "|" & _
                    sSheetCategory & "|" & _
                    sSheetHeader & "|"
                If sht.ListObjects.Count = 1 Then
                    sRowToWrite = sRowToWrite & sht.ListObjects(1).Name & "|" & sht.ListObjects(1).HeaderRowRange.Columns.Count
                End If
            End If
            On Error GoTo 0
                
            Print #iFileNo, sRowToWrite;
        End If
        
    Next sht
    
    Close #iFileNo

End Sub

