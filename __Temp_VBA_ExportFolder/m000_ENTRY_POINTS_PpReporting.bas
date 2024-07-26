Attribute VB_Name = "m000_ENTRY_POINTS_PpReporting"
'@Folder "SpreadsheetBI"
Option Explicit
Option Private Module

Public Const cPR_MaxStorageRecords As Long = 1000000  'PR  = PowerReport


Public Type TypeReportUserSelection
    SelectionMade As Boolean
    ReportNames() As String
    NumberOfSelectedReports As Long
End Type


Sub SaveReportMetadataInActiveWorkbook()
'Reads all report metadata from reports in active workbook and saves

    StandardEntry
    SaveReportingPowerPivotMetaData ActiveWorkbook
    InsertIndexPage ActiveWorkbook
    StandardExit
    
End Sub


Sub CreateReportFromMetadata()

    Dim vStorageObjReportStructure As Variant
    Dim vStorageObjQueriesForSelectedReports As Variant
    Dim UserReportSelection As TypeReportUserSelection
    Dim i As Long
    Dim PwrPvtReport As ReportingPowerPivot
    Dim sReportName As String
    Dim wkb As Workbook
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
    Set wkb = ActiveWorkbook

    With UserReportSelection

        If .SelectionMade = False Then GoTo Exitpoint

        For i = LBound(.ReportNames) To UBound(.ReportNames)
            sReportName = .ReportNames(i)
                Set PwrPvtReport = New ReportingPowerPivot
                PwrPvtReport.CreateEmptyPowerPivotReport wkb, sReportName
                DesignPowerPivotReportBasedOnStoredData _
                    vStorageObjReportStructure, PwrPvtReport
        Next i

    End With


Exitpoint:
    InsertIndexPage ActiveWorkbook
    StandardExit

End Sub


