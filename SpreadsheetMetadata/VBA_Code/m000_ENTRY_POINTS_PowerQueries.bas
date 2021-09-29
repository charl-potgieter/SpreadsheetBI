Attribute VB_Name = "m000_ENTRY_POINTS_PowerQueries"
Option Explicit


Sub ExportPowerQueriesInActiveWorkbookToFiles()

    Dim sFolderSelected As String

    sFolderSelected = GetFolder
    If NumberOfFilesInFolder(sFolderSelected) <> 0 Then
        MsgBox ("Please select an empty folder...exiting")
        Exit Sub
    End If

    ExportPowerQueriesToFiles sFolderSelected, ActiveWorkbook
    MsgBox ("Queries Exported")

End Sub


Sub ExportPowerQueriesInActiveWorkbookToConsolidatedFile()

    ExportPowerQueriesToConsolidatedFile ActiveWorkbook
    
    MsgBox ("Queries Exported")

End Sub


Sub ExportNonStandardPowerQueriesInActiveWorkbookToFiles()
'Exports power queries without fn_std or template_std prefix

    Dim sFolderSelected As String

    sFolderSelected = GetFolder
    If NumberOfFilesInFolder(sFolderSelected) <> 0 Then
        MsgBox ("Please select an empty folder...exiting")
        Exit Sub
    End If

    ExportNonStandardPowerQueriesToFiles sFolderSelected, ActiveWorkbook
    MsgBox ("Queries Exported")

End Sub


Sub ImportPowerQueriesFromSelectedFolderNonRecursive()

    Dim sFolderSelected As String

    sFolderSelected = GetFolder
    ImportOrRefreshPowerQueriesInFolder sFolderSelected, False
    MsgBox ("Queries imported")

End Sub



Sub ImportPowerQueriesFromSelectedFolderRecursive()

    Dim sFolderSelected As String

    sFolderSelected = GetFolder
    ImportOrRefreshPowerQueriesInFolder sFolderSelected, True
    MsgBox ("Queries imported")

End Sub



Sub ImportSelectedPowerQueries()
'Requires reference to Microsoft Scripting runtime library

    Dim sPowerQueryFilePath As String
    Dim sPowerQueryName As String
    Dim fDialog As FileDialog
    Dim fso As FileSystemObject
    Dim i As Integer

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    With fDialog
        .AllowMultiSelect = True
        .Title = "Select power query / queries"
        .InitialFileName = ThisWorkbook.Path
        .Filters.Clear
        .Filters.Add "m Power Query Files", "*.m"
    End With

    'fDialog.Show value of -1 below means success
    If fDialog.Show = -1 Then
        For i = 1 To fDialog.SelectedItems.Count
            sPowerQueryFilePath = fDialog.SelectedItems(i)
            Set fso = New FileSystemObject
            sPowerQueryName = Replace(fso.GetFileName(sPowerQueryFilePath), ".m", "")
            ImportOrRefreshSinglePowerQuery sPowerQueryFilePath, sPowerQueryName, ActiveWorkbook
        Next i
    End If

    MsgBox ("Queries imported")

End Sub




Sub CreatePowerQueryGeneratorSheet()
'Creates a sheet in active workbook to be utilsed for the generation of "hard coded" power query tables
'utilising sub GeneratePowerQueryTable

    Dim iMsgBoxResponse As Integer

    StandardEntry
    'Give user choice to delete sheet or cancel if sheet already exists?
    If SheetExists(ActiveWorkbook, "PqTableGenerator") Then
        iMsgBoxResponse = MsgBox("Sheet already exists, delete?", vbQuestion + vbYesNo + vbDefaultButton2)
        If iMsgBoxResponse = vbNo Then
            Exit Sub
        Else
            ActiveWorkbook.Sheets("PqTableGenerator").Delete
        End If
    End If

    CreateTableGeneratorSheet ActiveWorkbook
    StandardExit
End Sub


Sub GeneratePowerQueryTable()
'Generates a hard coded power query table from ActiveWorkbook.Sheets("TableGenerator").ListObjects("tbl_TableGenerator")
'Query name is as per defined name on sheet entitled "TableName"
'Column types are stored in cells 2 rows above the tbl_Generator table

    Dim sQueryName As String
    Dim sQueryText As String
    Dim i As Integer
    Dim j As Integer
    Dim lo As ListObject

    StandardEntry
    sQueryName = ActiveWorkbook.Sheets("PqTableGenerator").Range("TableName")

    If QueryExists(sQueryName, ActiveWorkbook) Then
        MsgBox ("Query with the same name already exists.  New query not generated")
        Exit Sub
    End If

    Set lo = ActiveWorkbook.Sheets("PqTableGenerator").ListObjects("tbl_PqTableGenerator")


    sQueryText = "let" & vbCrLf & vbCrLf & _
        "    tbl = Table.FromRecords({" & vbCrLf

    'Table from records portion of power query
    For i = 1 To lo.DataBodyRange.Rows.Count
        For j = 1 To lo.HeaderRowRange.Cells.Count
            If j = 1 Then
                sQueryText = sQueryText & "        ["
            End If
            sQueryText = sQueryText & _
                lo.HeaderRowRange.Cells(j) & " = """ & lo.ListColumns(j).DataBodyRange.Cells(i) & """"
            If j <> lo.HeaderRowRange.Cells.Count Then
                sQueryText = sQueryText & ", "
            ElseIf i <> lo.DataBodyRange.Rows.Count Then
                sQueryText = sQueryText & "], " & vbCrLf
            Else
                sQueryText = sQueryText & "]" & vbCrLf
            End If
        Next j
        If i = lo.DataBodyRange.Rows.Count Then
            sQueryText = sQueryText & "        }), " & vbCrLf & vbCrLf
        End If
    Next i


    'Changed Type portion of power query
    sQueryText = sQueryText & "    ChangedType = Table.TransformColumnTypes(" & vbCrLf & _
        "       tbl, " & vbCrLf & "        {" & vbCrLf

    For j = 1 To lo.HeaderRowRange.Cells.Count
        sQueryText = sQueryText & "            {""" & lo.HeaderRowRange.Cells(j) & """, " & lo.HeaderRowRange.Cells(j).Offset(-2, 0) & "}"
        If j <> lo.HeaderRowRange.Cells.Count Then
            sQueryText = sQueryText & "," & vbCrLf
        Else
            sQueryText = sQueryText & vbCrLf
        End If
    Next j
    sQueryText = sQueryText & vbCrLf & "        })" & vbCrLf & vbCrLf & _
        "in" & vbCrLf & "    ChangedType"

    ActiveWorkbook.Queries.Add sQueryName, sQueryText

    MsgBox ("Query Generated")
    StandardExit

End Sub


Sub CopyPowerQueriesFromWorkbook()
'Copies power queries from selected workbook into active workbook


    Dim fDialog As FileDialog, Result As Integer
    Dim sFilePathAndName As String
    Dim bWorkbookIsOpen As Boolean
    Dim fso As New FileSystemObject
    Dim sWorkbookName As String
    Dim wkbSource As Workbook
    Dim wkbTarget As Workbook
    Dim qry As WorkbookQuery

    StandardEntry

    Set wkbTarget = ActiveWorkbook
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)

    'Get file from file picker
    fDialog.AllowMultiSelect = False
    fDialog.InitialFileName = ActiveWorkbook.Path
    fDialog.Filters.Clear
    fDialog.Filters.Add "Excel files", "*.xlsx, *.xlsm"
    fDialog.Filters.Add "All files", "*.*"

    'Exit sub  if no file is selected
    If fDialog.Show <> -1 Then
       GoTo ExitPoint
    End If

    sFilePathAndName = fDialog.SelectedItems(1)
    sWorkbookName = fso.GetFileName(sFilePathAndName)


    If sWorkbookName = ActiveWorkbook.Name Then
        MsgBox ("Cannot copy between 2 workbooks with the same name, exiting...")
        GoTo ExitPoint
    End If


    'Open source workbook if not open
    If WorkbookIsOpen(sWorkbookName) Then
        bWorkbookIsOpen = True
        Set wkbSource = Workbooks(sWorkbookName)
    Else
        bWorkbookIsOpen = False
        Set wkbSource = Application.Workbooks.Open(sFilePathAndName)
    End If

    'Copy queries from source to active workbook
    For Each qry In wkbSource.Queries
        If QueryExists(qry.Name, wkbTarget) Then
            wkbTarget.Queries(qry.Name).Formula = qry.Formula
        Else
            wkbTarget.Queries.Add qry.Name, qry.Formula
        End If
    Next qry

    'Close source workbook if it was not open before this sub was run
    If Not bWorkbookIsOpen Then
        wkbSource.Close
    End If

    wkbTarget.Activate
    MsgBox ("Power Queries copied")

ExitPoint:
    StandardExit

End Sub

Sub CreateTextReferencedPowerQueriesInActiveWorkbook()

    Dim FilePaths() As String
    Dim i As Integer
    Dim fso As FileSystemObject
    Dim wkb As Workbook
    Dim QueryText As String
    Dim QueryName As String
    
    Set fso = New FileSystemObject
    Set wkb = ActiveWorkbook
    GetPowerQueryFileNamesFromUser FilePaths
    For i = LBound(FilePaths) To UBound(FilePaths)
        QueryName = fso.GetBaseName(FilePaths(i))
        QueryText = PowerQueryReferencedToTextFile(FilePaths(i))
        wkb.Queries.Add QueryName, QueryText
    Next i

    MsgBox ("Queries created")

End Sub


Sub TempDeleteAllPQ()
'Deletes all power queries in active workbook

    Dim qry As WorkbookQuery

    StandardEntry
    For Each qry In ThisWorkbook.Queries
        qry.Delete
    Next qry
    StandardExit

End Sub


