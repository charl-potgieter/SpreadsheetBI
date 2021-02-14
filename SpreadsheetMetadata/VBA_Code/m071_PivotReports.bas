Attribute VB_Name = "m071_PivotReports"
Option Explicit


Function ExtractPivotReportMetadataFromReports(ByRef wkb As Workbook) As TypePivotReport()

    Dim sht As Worksheet
    Dim pvt As PivotTable
    Dim pvtCubeField As cubeField
    Dim pvtField As PivotField
    Dim tempPvtField As PivotField
    Dim pvtReportDetails() As TypePivotReport
    Dim i As Integer
    Dim j As Integer
    Dim arrPvtTableProperties As Variant
    Dim arrPvtCubeFieldProperties As Variant
    Dim arrPvtFieldProperties As Variant
    Dim vPvtTableProperty As Variant
    Dim vPvtCubeFieldProperty As Variant
    Dim vPvtFieldProperty As Variant
    Dim sPropertyValue

    arrPvtTableProperties = Split(PivotTableProperties, "|")
    arrPvtCubeFieldProperties = Split(CubeFieldProperties, "|")
    arrPvtFieldProperties = Split(PivotFieldProperties, "|")
    
    i = 0
    
    
    For Each sht In wkb.Worksheets
        If sht.PivotTables.Count = 1 And _
            SheetLevelRangeNameExists(sht, "SheetHeading") And _
            SheetLevelRangeNameExists(sht, "SheetCategory") Then
                ReDim Preserve pvtReportDetails(i)
                
                Set pvt = sht.PivotTables(1)
                
                
                pvtReportDetails(i).SheetName = sht.Name
                With pvtReportDetails(i).ReportingSheet
                    .Name = sht.Name
                    Set .Properties = New Dictionary
                    .Properties.Add "SheetHeading", sht.Range("SheetHeading")
                    .Properties.Add "SheetCategory", sht.Range("SheetCategory")
                End With
                
                With pvtReportDetails(i).PvtTable
                    .Name = pvt.Name
                    Set .Properties = New Dictionary
                    For Each vPvtTableProperty In arrPvtTableProperties
                        sPropertyValue = CallByName(pvt, vPvtTableProperty, VbGet)
                        .Properties.Add vPvtTableProperty, sPropertyValue
                    Next vPvtTableProperty
                End With
                
                j = 0
                For Each pvtCubeField In pvt.CubeFields
                    If Not pvtCubeField.Orientation = xlHidden Then
                        ReDim Preserve pvtReportDetails(i).PvtCubeFields(j)
                        With pvtReportDetails(i).PvtCubeFields(j)
                            .Name = pvtCubeField.Name
                            'Some of Object model properties are set at cube field level
                            Set .Properties = New Dictionary
                            For Each vPvtCubeFieldProperty In arrPvtCubeFieldProperties
                                sPropertyValue = CallByName(pvtCubeField, vPvtCubeFieldProperty, VbGet)
                                .Properties.Add vPvtCubeFieldProperty, sPropertyValue
                            Next vPvtCubeFieldProperty
                        End With
                        j = j + 1
                    End If
                Next pvtCubeField
                
                
                'Some of Object model properties are set at on PivotField object rather than
                'cubefield object for some unknown reason
                j = 0
                For Each pvtField In pvt.PivotFields
                    ReDim Preserve pvtReportDetails(i).PvtFields(j)
                    With pvtReportDetails(i).PvtFields(j)
                        .Name = pvtField.Name
                        Set .Properties = New Dictionary
                        For Each vPvtFieldProperty In arrPvtFieldProperties
                            'Depending on field orientation the property may generate an error
                            'so ignore by setting to <NULL>
                            On Error Resume Next
                            
                            'Looks like a VBA bug? (or something I don't understand) requiring
                            'use of tempPvtField rather than simply just pvtField
                            Set tempPvtField = pvt.PivotFields(pvtField.Name)
                            sPropertyValue = CallByName(tempPvtField, vPvtFieldProperty, VbGet)
                            If Err.Number = 0 Then
                                .Properties.Add vPvtFieldProperty, sPropertyValue
                            Else
                                .Properties.Add vPvtFieldProperty, "<NULL>"
                            End If
                            On Error GoTo 0
                        Next vPvtFieldProperty
                    End With
                    j = j + 1
                Next pvtField
                
                i = i + 1
        End If
    Next sht

    ExtractPivotReportMetadataFromReports = pvtReportDetails
    
End Function



Function CreateReportSheet(ByRef wkb As Workbook, ByRef pvtReportMetaData As TypePivotReport) As Worksheet

    Dim sSheetName As String
    Dim sht As Worksheet
    
    
    sSheetName = pvtReportMetaData.SheetName
    If SheetExists(wkb, sSheetName) Then
        wkb.Sheets(sSheetName).Delete
    End If

    Set sht = wkb.Sheets.Add(After:=ActiveSheet)

    sht.Name = sSheetName
    Set CreateReportSheet = sht
    


End Function



Function CreateEmptyPowerPivotTable(ByRef sht As Worksheet) As PivotTable

    Dim pvt As PivotTable

    'Create pivot table from data model
    'report is subsequently shifted down once design is complete and size is known
    Set pvt = ActiveWorkbook.PivotCaches.Create(SourceType:=xlExternal, SourceData:= _
        ActiveWorkbook.Connections("ThisWorkbookDataModel"), Version:=6). _
        CreatePivotTable(sht.Range("B1"))

    Set CreateEmptyPowerPivotTable = pvt


End Function


Sub SetPivotTableProperties(ByRef pvt As PivotTable, ByRef pvtReportMetaData As TypePivotReport)

    Dim property
    Dim sValue As String

    pvt.Name = pvtReportMetaData.PvtTable.Name
    
    With pvtReportMetaData.PvtTable.Properties
        For Each property In .Keys
            On Error Resume Next
            CallByName pvt, property, VbLet, .item(property)
            If Err.Number <> 0 Then
                Debug.Print property
            End If
            On Error GoTo 0
        Next property
    End With

End Sub

Sub SetPivotCubeFieldsProperties(ByRef pvt As PivotTable, ByRef pvtReportMetaData As TypePivotReport)

    Dim pvtCubeField As cubeField
    Dim property
    Dim sValue As String
    Dim i As Double


    For i = LBound(pvtReportMetaData.PvtCubeFields) To UBound(pvtReportMetaData.PvtCubeFields)
        Set pvtCubeField = pvt.CubeFields(pvtReportMetaData.PvtCubeFields(i).Name)
        pvtCubeField.Orientation = pvtReportMetaData.PvtCubeFields(i).Properties.item("Orientation")
        pvtCubeField.Caption = pvtReportMetaData.PvtCubeFields(i).Properties.item("Caption")
    Next i

    'Can only set position after cubefield orientation has been set on pivot table
    For i = LBound(pvtReportMetaData.PvtCubeFields) To UBound(pvtReportMetaData.PvtCubeFields)
        Set pvtCubeField = pvt.CubeFields(pvtReportMetaData.PvtCubeFields(i).Name)
        
        'Setting to the same number seems to make field disappear for at least value field
        'where position is 1.
        If pvtCubeField.Position <> pvtReportMetaData.PvtCubeFields(i).Properties.item("Position") Then
            pvtCubeField.Position = pvtReportMetaData.PvtCubeFields(i).Properties.item("Position")
        End If
    Next i


End Sub




Sub SetPivotFieldsProperties(ByRef pvt As PivotTable, ByRef pvtReportMetaData As TypePivotReport)

    Dim pvtField As PivotField
    Dim property
    Dim sValue As String
    Dim i As Double


    For i = LBound(pvtReportMetaData.PvtFields) To UBound(pvtReportMetaData.PvtFields)



        With pvtReportMetaData.PvtFields(i).Properties
            For Each property In .Keys
                Set pvtField = pvt.PivotFields(pvtReportMetaData.PvtFields(i).Name)
                'Neeed to ignore errors as not all properties will apply depending on field
                'orientation
                On Error Resume Next
                CallByName pvtField, property, VbLet, .item(property)
                On Error GoTo 0
            Next property
        End With

    Next i

End Sub


Sub FormatPivotReportSheet(ByRef sht As Worksheet, ByRef pvtReportMetaData As TypePivotReport)

    sht.Range("1:6").Insert (xlShiftDown)
    FormatSheet sht
    sht.Range("SheetHeading") = pvtReportMetaData.ReportingSheet.Properties.item("SheetHeading")
    sht.Range("SheetCategory") = pvtReportMetaData.ReportingSheet.Properties.item("SheetCategory")


End Sub
