Attribute VB_Name = "m500_Core_PpReporting"
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
            ReadSelectedReportsFromComboBox uf.lbReports, .ReportNames
            .NumberOfSelectedReports = UBound(.ReportNames) - LBound(.ReportNames)
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
    WriteReportData vStorageObject, sReportName, _
        "Sheet", PwrPvtReport.SheetProperties
    WriteReportData vStorageObject, sReportName, _
        "PivotTable", PwrPvtReport.PivotTableProperties
    WriteReportData vStorageObject, sReportName, _
        "ColumnWidths", PwrPvtReport.ColumnWidths
    WriteReportData vStorageObject, sReportName, "CubeFields", _
            PwrPvtReport.CubeFieldProperties
    WriteReportData vStorageObject, sReportName, "PivotFields", _
            PwrPvtReport.PivotFieldProperties


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




'
'Function AssignReportWorkbook(ByVal wkbCurrent As Workbook) As Workbook
''Assign workbook based on whether user selects to save in new workbook or not.
'
'    Dim sFilePath As String
'    Dim sFileName As String
'    Dim sFilePathAndName As String
'
'    If bSaveInNewWorkbook Then
'        sFilePath = GetReportFilePath(wkbCurrent)
'        If Not FolderExists(sFilePath) Then
'            CreateFolder sFilePath
'        End If
'        sFileName = GetReportFileName(wkbCurrent)
'        sFilePathAndName = sFilePath & Application.PathSeparator & sFileName
'        ActiveWorkbook.SaveCopyAs (sFilePathAndName)
'        Set AssignReportWorkbook = Workbooks.Open(sFilePathAndName)
'    Else
'        Set AssignReportWorkbook = wkbCurrent
'    End If
'
'End Function



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
    ByVal wkb As Workbook, ByRef ReportNames() As String)

    Dim sQueriesRequired As Variant
    Dim i As Long
    Dim con As WorkbookConnection
    Dim bDeleteQuery As Boolean
    Dim sTableName

    sQueriesRequired = ReadQueriesForReportList(vStorageObject, ReportNames)

    'Do not delete any tables if none are listed as being required (default to retain all)
    If IsNull(sQueriesRequired) Then
        Exit Sub
    End If

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

    Set vStorageObject = AssignReportStructureStorage(ActiveWorkbook)
    For Each sht In ActiveWorkbook.Worksheets
        Set PwrPvtReport = New ReportingPowerPivot
        bValidAssignment = PwrPvtReport.AssignToExistingSheet(sht)
        If bValidAssignment Then
            SaveSinglePowerPivotReportDataToStorage vStorageObject, PwrPvtReport
        End If
        Set PwrPvtReport = Nothing
    Next sht

End Sub



Sub ReadSelectedReportsFromComboBox(ByVal lb As MSForms.ListBox, ByRef ReportNames() As String)
'Reads from the combo box into a base zero one-dimensional array

    Dim CurrentListItem As Long
    Dim SelectedItemsCounter As Long
    Const cCombBoxReportNameColumn As Integer = 0
    Const cComboBoxReportTypeColumn As Integer = 1

    'Leave array empty if nothing selected
    If Not UserFormListBoxHasSelectedItems(lb) Then Exit Sub

    SelectedItemsCounter = 0
    For CurrentListItem = 0 To lb.ListCount - 1
        If lb.Selected(CurrentListItem) = True Then
            ReDim Preserve ReportNames(SelectedItemsCounter)
            ReportNames(SelectedItemsCounter) = _
                lb.List(CurrentListItem, cCombBoxReportNameColumn)
            SelectedItemsCounter = SelectedItemsCounter + 1
        End If
    Next CurrentListItem

End Sub



