Attribute VB_Name = "ReportingGeneral"
'@Folder("Reporting")
Option Explicit


Function GetUserPivotReportSelection() As TypePivotReportUserSelection

    Dim uf As ufPivotReportGenerator
    Dim UserSelection As TypePivotReportUserSelection

    Set uf = New ufPivotReportGenerator
    uf.Show
    
    With UserSelection
        If uf.bCancelled Or Not UserFormListBoxHasSelectedItems(uf.lbReports) Then
            .SelectionMade = False
        Else
            .SelectionMade = True
            .ReportList = UserFormListBoxSelectedArray(uf.lbReports)
            .RetainLiveInCurrentWorkbook = uf.chkRetainLiveReport.Value
            .GenerateValueCopyNewWorkbook = uf.chkValueCopy.Value
            Select Case True
                Case uf.obPowePivotSource.Value = True
                    .PivotReportType = PowerPivotSource
                Case uf.obExcelTableSource.Value = True
                    .PivotReportType = ExcelTableSource
                Case uf.obExcelTableOnly.Value = True
                    .PivotReportType = ExcelTableOnly
            End Select
        End If
    End With
    
    Unload uf
    Set uf = Nothing
    GetUserPivotReportSelection = UserSelection

End Function


Public Sub SaveSinglePivotReportDataToStorage(ByVal vStorageObject As Variant, _
    ByVal PvtReport As PivotReport)
    
    Dim PvtCubeField As CubeField
    Dim pvtField As PivotField
    Dim sReportName As String
    
    sReportName = PvtReport.ReportName
    DeleteExistingPivotReportData vStorageObject, sReportName
    WritePivotReportData vStorageObject, sReportName, "Sheet", PvtReport.SheetProperties
    WritePivotReportData vStorageObject, sReportName, "PivotTable", PvtReport.PivotTableProperties
    WritePivotReportData vStorageObject, sReportName, "ColumnWidths", PvtReport.ColumnWidths
    WritePivotReportData vStorageObject, sReportName, "CubeFields", _
            PvtReport.CubeFieldProperties
    WritePivotReportData vStorageObject, sReportName, "PivotFields", _
            PvtReport.PivotFieldProperties
    

End Sub


Public Sub DesignPivotReportBasedOnStoredData(ByVal vStorageObject As Variant, _
    ByVal PvtReport As PivotReport)
'Reads from Storage structure into the pivot table

'TODO complete ReadSinglePivotReportData

    PvtReport.PivotTableProperties = ReadPivotReportProperties(vStorageObject, PvtReport, _
        "PivotTable")
    PvtReport.CubeFieldProperties = ReadPivotReportProperties(vStorageObject, PvtReport, _
        "CubeFields")
    PvtReport.PivotFieldProperties = ReadPivotReportProperties(vStorageObject, PvtReport, _
        "PivotFields")
    PvtReport.SheetProperties = ReadPivotReportProperties(vStorageObject, PvtReport, _
        "Sheet")
    
End Sub

