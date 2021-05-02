Attribute VB_Name = "Reporting_General"
'@Folder "Storage.Reporting"
Option Explicit


Function GetUserReportSelection() As TypeReportUserSelection

    Dim uf As ufPivotReportGenerator
    Dim UserSelection As TypeReportUserSelection

    Set uf = New ufPivotReportGenerator
    uf.Show
    
    With UserSelection
        If uf.bCancelled Or Not UserFormListBoxHasSelectedItems(uf.lbReports) Then
            .SelectionMade = False
        Else
            .SelectionMade = True
            .ReportList = UserFormListBoxSelectedArray(uf.lbReports)
            .SaveInNewWorkbook = uf.chkSaveInNewSpreadsheet
            Select Case True
                Case uf.obPowerPivotSource.Value = True
                    .ReportType = PowerPivotSource
                Case uf.obExcelTableSource.Value = True
                    .ReportType = ExcelTableSource
                Case uf.obExcelTableOnly.Value = True
                    .ReportType = ExcelTableOnly
            End Select
        End If
    End With
    
    Unload uf
    Set uf = Nothing
    GetUserReportSelection = UserSelection

End Function


Public Sub SaveSinglePowerPivotReportDataToStorage(ByVal vStorageObject As Variant, _
    ByVal PwrPvtReport As ReportingPowerPivot)
    
    Dim sReportName As String
    
    sReportName = PwrPvtReport.Name
    DeleteExistingReportData vStorageObject, sReportName
    WriteReportData vStorageObject, sReportName, "Sheet", PwrPvtReport.SheetProperties
    WriteReportData vStorageObject, sReportName, "PivotTable", PwrPvtReport.PivotTableProperties
    WriteReportData vStorageObject, sReportName, "ColumnWidths", PwrPvtReport.ColumnWidths
    WriteReportData vStorageObject, sReportName, "CubeFields", _
            PwrPvtReport.CubeFieldProperties
    WriteReportData vStorageObject, sReportName, "PivotFields", _
            PwrPvtReport.PivotFieldProperties
    

End Sub


Public Sub SaveSingleTableReportDataToStorage(ByVal vStorageObject As Variant, _
    ByVal RptTable As ReportingTable)

    Dim sReportName As String
    
    sReportName = RptTable.Name
    DeleteExistingReportData vStorageObject, sReportName
    WriteReportData vStorageObject, sReportName, "Sheet", RptTable.SheetProperties
    WriteReportData vStorageObject, sReportName, "Formulas", RptTable.Formulas
    WriteReportData vStorageObject, sReportName, "NumberFormats", RptTable.NumberFormatting

 End Sub




Public Sub DesignPowerPivotReportBasedOnStoredData(ByVal vStorageObject As Variant, _
    ByVal PwrPvtReport As ReportingPowerPivot)
'Reads from Storage structure into the power pivot

    PwrPvtReport.PivotTableProperties = ReadReportProperties(vStorageObject, _
        PwrPvtReport.Name, "PivotTable")
    PwrPvtReport.CubeFieldProperties = ReadReportProperties(vStorageObject, _
        PwrPvtReport.Name, "CubeFields")
    PwrPvtReport.PivotFieldProperties = ReadReportProperties(vStorageObject, _
        PwrPvtReport.Name, "PivotFields")
    PwrPvtReport.SheetProperties = ReadReportProperties(vStorageObject, _
        PwrPvtReport.Name, "Sheet")
    'Columnwidth needs to be set last
    PwrPvtReport.ColumnWidths = ReadReportProperties(vStorageObject, _
        PwrPvtReport.Name, "ColumnWidths")
    
End Sub


Public Sub DesignPowerTableReportBasedOnStoredData(ByVal vStorageObject As Variant, _
    ByVal TableReport As ReportingTable, ByVal sDaxTableQueryPath As String)
'Reads from Storage structure into the reporting table

    Dim sDaxQuery As String

    'TODO set other properties here.
    TableReport.SheetProperties = ReadReportProperties(vStorageObject, _
        TableReport.Name, "Sheet")
    sDaxQuery = m030_FileUtilities.ReadTextFileIntoString(sDaxTableQueryPath)
    TableReport.DaxQuery = sDaxQuery

End Sub


Function AssignReportWorkbook(ByVal wkbCurrent As Workbook, _
    ByVal bSaveInNewWorkbook As Boolean) As Workbook
'Assign workbook based on whether user selects to save in new workbook or not.

    Dim sFilePath As String
    Dim sFileName As String
    Dim sFilePathAndName As String

    If bSaveInNewWorkbook Then
        sFilePath = GetReportFilePath(wkbCurrent)
        If Not FolderExists(sFilePath) Then
            CreateFolder sFilePath
        End If
        sFileName = GetReportFileName(wkbCurrent)
        sFilePathAndName = sFilePath & Application.PathSeparator & sFileName
        ActiveWorkbook.SaveCopyAs (sFilePathAndName)
        Set AssignReportWorkbook = Workbooks.Open(sFilePathAndName)
    Else
        Set AssignReportWorkbook = wkbCurrent
    End If

End Function



Function GetReportFilePath(ByVal wkb As Workbook) As String
'Returns file path when report is saved in a new workbook

    GetReportFilePath = ActiveWorkbook.Path & Application.PathSeparator & "ReportsGenerated"
        
End Function



Function GetReportFileName(ByVal wkb As Workbook) As String
'Returns file name when report is saved in a new workbook

    Dim fso As FileSystemObject
    Dim sWkbFileName As String
    Dim sWkbExtension As String
    
    sWkbFileName = Year(Now) & Month(Now) & Day(Now) & "_" & Hour(Now) & _
        Minute(Now) & Second(Now) & "_ReportGenerated"
    Set fso = CreateObject("Scripting.FileSystemObject")
    sWkbExtension = fso.GetExtensionName(wkb.Name)
    
    GetReportFileName = sWkbFileName & "." & sWkbExtension
        
End Function


Sub DeleteNonReportSheets(ByVal wkb As Workbook, ByRef ReportNames() As String)
'Deletes surplus sheets when reports are generated to a new workbook
    
    Dim sht As Worksheet
    Dim i As Long
    Dim ReportSht As ReportingSheet
    Dim bReportSheetAssigned As Boolean
    Dim bRetainSheet As Boolean
    

    For Each sht In wkb.Sheets
        bRetainSheet = False
        Set ReportSht = New ReportingSheet
        bReportSheetAssigned = ReportSht.AssignExistingSheet(sht)
        If bReportSheetAssigned Then
            i = LBound(ReportNames)
            Do While i <= UBound(ReportNames) And Not (bRetainSheet)
                If ReportSht.Heading = ReportNames(i) Then bRetainSheet = True
                i = i + 1
            Loop
            If Not bRetainSheet Then ReportSht.Delete
        Else
            'Not a valid reportsheet - can safely delete
            sht.Delete
        End If
    Next sht

End Sub


Sub DeleteUnusedDataModelTables(ByVal vStorageObject As Variant, _
    ByVal wkb As Workbook, ByRef sReportNames() As String)
    
    Dim sQueriesRequired As Variant
    Dim i As Long
    Dim con As WorkbookConnection
    Dim bDeleteQuery As Boolean
    Dim sTableName
    
    sQueriesRequired = ReadQueriesForReportList(vStorageObject, sReportNames)

    For Each con In wkb.Connections
        bDeleteQuery = True
        If con.Type = xlConnectionTypeOLEDB And con.InModel Then
            i = LBound(sQueriesRequired)
            Do While i <= UBound(sQueriesRequired) And bDeleteQuery
                sTableName = Replace(con.OLEDBConnection.CommandText, """", "")
                If sTableName = sQueriesRequired(i) Then
                    bDeleteQuery = False
                End If
                i = i + 1
            Loop
            If bDeleteQuery Then
                con.Delete
                If QueryExists(sTableName, wkb) Then wkb.Queries(sTableName).Delete
            End If
        End If
    Next con
            

End Sub


Sub SaveReportingPowerPivotMetaData(ByVal wkb As Workbook)

    Dim sht As Worksheet
    Dim bValidAssignment As Boolean
    Dim PwrPvtReport As ReportingPowerPivot
    Dim vStorageObject As Variant 'abstract away the storage structure
    
    Set vStorageObject = AssignPivotReportStructureStorage(ActiveWorkbook)
    For Each sht In ActiveWorkbook.Worksheets
        Set PwrPvtReport = New ReportingPowerPivot
        bValidAssignment = PwrPvtReport.AssignToExistingSheet(sht)
        If bValidAssignment Then
            SaveSinglePowerPivotReportDataToStorage vStorageObject, PwrPvtReport
        End If
        Set PwrPvtReport = Nothing
    Next sht

End Sub


Sub SaveReportingTableMetadata(ByVal wkb As Workbook)

    Dim sht As Worksheet
    Dim bValidAssignment As Boolean
    Dim RptTable As ReportingTable
    Dim vStorageObject As Variant 'abstract away the storage structure
    Dim sDaxQueryFilePath As String
    Dim sDaxQueryFilePathAndFileName As String
    Const csSubDirectory As String = "DaxTableQueries"
    
    Set vStorageObject = AssignTableReportStorage(ActiveWorkbook)
    For Each sht In ActiveWorkbook.Worksheets
        Set RptTable = New ReportingTable
        bValidAssignment = RptTable.AssignToExistingSheet(sht)
        If bValidAssignment Then
            SaveSingleTableReportDataToStorage vStorageObject, RptTable
            sDaxQueryFilePath = ThisWorkbook.Path & Application.PathSeparator & _
                csSubDirectory
            If Not FolderExists(sDaxQueryFilePath) Then CreateFolder (sDaxQueryFilePath)
            sDaxQueryFilePathAndFileName = sDaxQueryFilePath & Application.PathSeparator & _
                 RptTable.Name & ".Dax"
            WriteStringToTextFile RptTable.DaxQuery, sDaxQueryFilePathAndFileName
        End If
        Set RptTable = Nothing
    Next sht

End Sub

