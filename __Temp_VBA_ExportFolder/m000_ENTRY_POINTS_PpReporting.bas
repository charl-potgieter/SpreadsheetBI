Attribute VB_Name = "m000_ENTRY_POINTS_PpReporting"
Option Explicit

Public Enum EnumReportType
    PowerPivotSource
    ExcelTableOnly
End Enum

Public Type TypeReportRecord
    ReportName As String
    ReportType As EnumReportType
End Type

Public Type TypeReportUserSelection
    SelectionMade As Boolean
    ReportList() As TypeReportRecord
    NumberOfSelectedReports As Long
    SaveInNewWorkbook As Boolean
    GenerateIndex As Boolean
    NumberOfReportsForIndexGeneration As Integer
End Type


Sub SaveReportMetadataInActiveWorkbook()
'Reads all report metadata from reports in active workbook and saves

    StandardEntry
    SaveReportingPowerPivotMetaData ActiveWorkbook
    SaveReportingTableMetadata ActiveWorkbook
    StandardExit

End Sub


Sub CreateReportFromMetadata()

    Dim vStorageObjReportStructure As Variant
    Dim vStorageObjQueriesForSelectedReports As Variant
    Dim UserReportSelection As TypeReportUserSelection
    Dim i As Long
    Dim PwrPvtReport As ReportingPowerPivot
    Dim TableReport As ReportingTable
    Dim sReportName As String
    Dim wkbSource As Workbook
    Dim wkbTarget As Workbook
    Dim sDaxTableQueryPath As String
    Const csSubDirectory As String = "DaxTableQueries"

    StandardEntry

    Set vStorageObjReportStructure = AssignReportStructureStorage(ActiveWorkbook, False)
    Set vStorageObjQueriesForSelectedReports = AssignPivotTableQueriesPerReport(ActiveWorkbook, False)

    'Exit if no report metadata exists on active sheet
    If vStorageObjReportStructure Is Nothing Then
        MsgBox ("No report metadata exists on active sheet")
        GoTo Exitpoint
    End If

    UserReportSelection = GetUserReportSelection
    Set wkbSource = ActiveWorkbook
    Set wkbTarget = AssignReportWorkbook(wkbSource, UserReportSelection.SaveInNewWorkbook)

    With UserReportSelection

        If .SelectionMade = False Then GoTo Exitpoint

        For i = LBound(.ReportList) To UBound(.ReportList)
            sReportName = .ReportList(i).ReportName
            Select Case .ReportList(i).ReportType
                Case PowerPivotSource
                    Set PwrPvtReport = New ReportingPowerPivot
                    PwrPvtReport.CreateEmptyPowerPivotReport wkbTarget, sReportName
                    DesignPowerPivotReportBasedOnStoredData _
                        vStorageObjReportStructure, PwrPvtReport
                Case ExcelTableOnly
                    sDaxTableQueryPath = wkbSource.Path & Application.PathSeparator & _
                        csSubDirectory & Application.PathSeparator & sReportName & ".dax"
                    Set TableReport = New ReportingTable
                    TableReport.CreateEmptyReportingTable wkbTarget, sReportName
                    DesignPowerTableReportBasedOnStoredData vStorageObjReportStructure, _
                        TableReport, sDaxTableQueryPath
                    TableReport.ApplyColourFormatting
            End Select
        Next i

        If .SaveInNewWorkbook Then
            DeleteNonReportSheets wkbTarget, .ReportList
            If Not vStorageObjQueriesForSelectedReports Is Nothing Then
                DeleteUnusedDataModelTables vStorageObjQueriesForSelectedReports, wkbTarget, .ReportList
            End If
            If .GenerateIndex And (UBound(.ReportList) - LBound(.ReportList) + 1) _
                >= .NumberOfReportsForIndexGeneration Then
                    InsertIndexPage wkbTarget
            End If

            wkbTarget.Save
        End If

    End With

Exitpoint:
    StandardExit

End Sub

