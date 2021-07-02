Attribute VB_Name = "TableLooper"
Option Explicit
Option Private Module


Function IsTableLooperSheet(ByVal sht As Worksheet) As Boolean
'Returns true if sheet contains
' - one listobject
' - contains one sheet level range name where name contains the word Selected
' - above range name contains list based data validation


    Dim nm As Name
    Dim iNumberOfNamesWithSelectedInName As Integer

    IsTableLooperSheet = False
    If sht.ListObjects.Count = 1 Then
        For Each nm In sht.Names
            iNumberOfNamesWithSelectedInName = 0
            If InStr(UCase(nm.Name), "SELECTED") <> 0 Then
                iNumberOfNamesWithSelectedInName = iNumberOfNamesWithSelectedInName + 1
            End If
        Next nm
        If iNumberOfNamesWithSelectedInName = 1 Then
            For Each nm In sht.Names
                If InStr(UCase(nm.Name), "SELECTED") <> 0 Then
                    On Error Resume Next
                    IsTableLooperSheet = (nm.RefersToRange.Validation.Type = xlValidateList)
                    On Error GoTo 0
                End If
            Next nm
        End If
    End If

End Function


Function GetLooperSelectionCell(ByVal sht As Worksheet) As Range

    Dim nm As Name

    For Each nm In sht.Names
        If InStr(UCase(nm.Name), "SELECTED") <> 0 Then
            Set GetLooperSelectionCell = nm.RefersToRange
        End If
    Next nm

End Function


Function InsertConsolLooperSheet(ByVal ReportSheetSource As ReportingSheet) As ReportingSheet

    Dim wkb As Workbook
    
    Set wkb = ReportSheetSource.Sheet.Parent

    On Error Resume Next
    wkb.Sheets("Consol_" & ReportSheetSource.Name).Delete
    Set InsertConsolLooperSheet = New ReportingSheet
    InsertConsolLooperSheet.Create ActiveWorkbook, ReportSheetSource.Sheet.Index
    InsertConsolLooperSheet.Name = "Consol_" & ReportSheetSource.Name
    InsertConsolLooperSheet.Category = ReportSheetSource.Category
    InsertConsolLooperSheet.Heading = "Consolidated " & ReportSheetSource.Heading
    On Error GoTo 0

End Function


Sub LoopSourceAndCopyToConsolSheet(ByVal ReportSheetSource As ReportingSheet, ByVal ReportSheetConsol As ReportingSheet)

    Dim loSource As ListObject
    Dim loTarget As ListObject
    Dim SelectionCell As Range
    Dim ValidationItems As Variant
    Dim rngStartOfConsolTable As Range
    Dim rngPasteTarget As Range
    Dim i As Long
    
    Set loSource = ReportSheetSource.Sheet.ListObjects(1)
    Set SelectionCell = GetLooperSelectionCell(ReportSheetSource.Sheet)
    Set rngStartOfConsolTable = ReportSheetConsol.Sheet.Range("B5")
    ValidationItems = GetDataValidationFromRangeReference(SelectionCell)
    
    For i = LBound(ValidationItems) To UBound(ValidationItems)
        SelectionCell.Value = ValidationItems(i)
        WaitForCalc
        If i = 0 Then
            loSource.Range.Copy
            rngStartOfConsolTable.PasteSpecial Paste:=xlPasteValues
            rngStartOfConsolTable.PasteSpecial Paste:=xlPasteFormats
        Else
            Set rngPasteTarget = rngStartOfConsolTable.CurrentRegion _
                .Offset(rngStartOfConsolTable.CurrentRegion.rows.Count, 0).Resize(1, 1)
            loSource.DataBodyRange.Copy
            rngPasteTarget.PasteSpecial Paste:=xlPasteValues
            rngPasteTarget.PasteSpecial Paste:=xlPasteFormats
        End If
    Next i

    Set loTarget = ReportSheetConsol.Sheet.ListObjects.Add(xlSrcRange, rngStartOfConsolTable.CurrentRegion, , xlYes)
    loTarget.Name = loSource.Name & "_Consolidated"

End Sub




Sub FilterOutExcludedItems(ByVal ReportSheetConsol As ReportingSheet)

    Dim lo As ListObject
    Dim bExclusionFieldExists As Boolean
    Dim lExclusionFieldIndex As Long
    Const csExclusionFieldName As String = "Include in consolidation"
        
    Set lo = ReportSheetConsol.Sheet.ListObjects(1)
    bExclusionFieldExists = WorksheetFunction.CountIfs(lo.HeaderRowRange, csExclusionFieldName) <> 0
    If bExclusionFieldExists Then
        lExclusionFieldIndex = WorksheetFunction.Match(csExclusionFieldName, lo.HeaderRowRange, 0)
        lo.Range.AutoFilter field:=lExclusionFieldIndex, Criteria1:="FALSE"
        
        On Error Resume Next 'In case there is nothing to delete
        lo.DataBodyRange.SpecialCells(xlCellTypeVisible).EntireRow.Delete
        On Error GoTo 0
        
        lo.Parent.ShowAllData
    End If

End Sub


Sub SetLoopTableAndSheetFormat(ByVal ReportSheetSource As ReportingSheet, ByVal ReportSheetConsol As ReportingSheet)

    Dim loSource As ListObject
    Dim loConsol As ListObject
    Dim i As Integer
    
    Set loSource = ReportSheetSource.Sheet.ListObjects(1)
    Set loConsol = ReportSheetConsol.Sheet.ListObjects(1)
    
    loConsol.HeaderRowRange.RowHeight = loSource.HeaderRowRange.RowHeight
    
    For i = 1 To loSource.ListColumns.Count
        loConsol.ListColumns(i).DataBodyRange.EntireColumn.ColumnWidth = loSource.ListColumns(i).DataBodyRange.EntireColumn.ColumnWidth
    Next i

    FormatTable loConsol
    
    ReportSheetConsol.Sheet.Activate
    ActiveWindow.SplitRow = 5
    ActiveWindow.FreezePanes = True

End Sub
