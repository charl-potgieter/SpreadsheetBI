Attribute VB_Name = "m060_CreateSpreadsheetMetaData"
Option Explicit
Option Private Module

Sub GenerateMetadataFileWorksheets(ByRef wkb As Workbook)



    Dim sht As Worksheet
    Dim sSheetHeader As String
    Dim sSheetCategory As String
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer
    Dim sFilePathAndName As String

    Set wkb = ActiveWorkbook
    sFolderPath = wkb.Path & Application.PathSeparator & "SpreadsheetMetadata"
    sFilePathAndName = sFolderPath & Application.PathSeparator & "MetadataWorksheets.txt"

    If Not (FolderExists(sFolderPath)) Then
        CreateFolder (sFolderPath)
    End If
    

    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "Name|Sheet Category|Sheet Header|Table Name|Number Of Table Columns|Table top Left Cells";
    
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
                    sRowToWrite = sRowToWrite & _
                        sht.ListObjects(1).Name & "|" & _
                        sht.ListObjects(1).HeaderRowRange.Columns.Count & "|" & _
                        sht.ListObjects(1).HeaderRowRange.Cells(1).Address
                Else
                    sRowToWrite = sRowToWrite & "||"
                End If
            End If
            On Error GoTo 0
                
            Print #iFileNo, sRowToWrite;
        End If
        
    Next sht
    
    Close #iFileNo

End Sub



Sub GenerateMetadataFileListObjectFields(ByRef wkb As Workbook)


    Dim i As Long
    Dim sht As Worksheet
    Dim lo As ListObject
    Dim sSheetHeader As String
    Dim sSheetCategory As String
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer
    Dim sFilePathAndName As String

    Set wkb = ActiveWorkbook
    sFolderPath = wkb.Path & Application.PathSeparator & "SpreadsheetMetadata"
    sFilePathAndName = sFolderPath & Application.PathSeparator & "ListObjectFields.txt"

    If Not (FolderExists(sFolderPath)) Then
        CreateFolder (sFolderPath)
    End If
    

    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "ListObjectName|ListObjectHeader|IsFormula|Formula";
    
    For Each sht In wkb.Worksheets
        
        If sht.Name <> "Index" Then
            On Error Resume Next 'Only write metadata if sheet meets below criteria
            sSheetHeader = sht.Names("SheetHeading").RefersToRange.Value
            sSheetCategory = sht.Names("SheetCategory").RefersToRange.Value
            Set lo = sht.ListObjects(1)
            
            If Err.Number = 0 And sht.ListObjects.Count = 1 Then
                For i = 1 To lo.HeaderRowRange.Columns.Count
                
                    sRowToWrite = vbCr & _
                        lo.Name & "|" _
                        & lo.HeaderRowRange.Cells(i) & "|" & _
                        lo.ListColumns(i).DataBodyRange.Cells(1).HasFormula & "|"
                        
                    If lo.ListColumns(i).DataBodyRange.Cells(1).HasFormula Then
                        sRowToWrite = sRowToWrite & lo.ListColumns(i).DataBodyRange.Cells(1).Formula
                    End If
                        
                    Print #iFileNo, sRowToWrite;
                Next i
            End If
            On Error GoTo 0
        End If
                
    Next sht
    Close #iFileNo

End Sub








Sub GenerateMetadataFileListObjectValues(ByRef wkb As Workbook)


    Dim i As Long
    Dim j As Long
    Dim sht As Worksheet
    Dim lo As ListObject
    Dim sSheetHeader As String
    Dim sSheetCategory As String
    Dim sRowToWrite As String
    Dim sFolderPath As String
    Dim iFileNo As Integer
    Dim sFilePathAndName As String

    Set wkb = ActiveWorkbook
    sFolderPath = wkb.Path & Application.PathSeparator & "SpreadsheetMetadata"
    sFilePathAndName = sFolderPath & Application.PathSeparator & "ListObjectFieldValues.txt"

    If Not (FolderExists(sFolderPath)) Then
        CreateFolder (sFolderPath)
    End If
    

    sRowToWrite = ""
    iFileNo = FreeFile 'Get first free file number
    Open sFilePathAndName For Output As #iFileNo
    
    'Write headers
    Print #iFileNo, "ListObjectName|ListObjectHeader|Value";
    
    For Each sht In wkb.Worksheets
        
        If sht.Name <> "Index" Then
            On Error Resume Next 'Only write metadata if sheet meets below criteria
            sSheetHeader = sht.Names("SheetHeading").RefersToRange.Value
            sSheetCategory = sht.Names("SheetCategory").RefersToRange.Value
            Set lo = sht.ListObjects(1)
            If Err.Number = 0 And sht.ListObjects.Count = 1 Then
                For i = 1 To lo.HeaderRowRange.Columns.Count
                    If Not (lo.ListColumns(i).DataBodyRange.Cells(1).HasFormula) Then
                        For j = 1 To lo.DataBodyRange.Rows.Count
                            sRowToWrite = vbCr & _
                                lo.Name & "|" & _
                                lo.ListColumns(i).Name & "|" & _
                                lo.ListColumns(i).DataBodyRange.Cells(j).Value
                                Print #iFileNo, sRowToWrite;
                        Next j
                    End If
                Next i
            End If
            On Error GoTo 0
        End If
                
    Next sht
    Close #iFileNo

End Sub




