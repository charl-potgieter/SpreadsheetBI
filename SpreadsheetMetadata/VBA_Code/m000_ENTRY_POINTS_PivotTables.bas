Attribute VB_Name = "m000_ENTRY_POINTS_PivotTables"
Option Explicit


Sub PivotTableFlatten()

    Dim pvt As PivotTable
    Dim pvtField As PivotField
    Dim b_mu As Boolean

    StandardEntry
    On Error Resume Next
    Set pvt = ActiveCell.PivotTable
    On Error GoTo 0

    If Not pvt Is Nothing Then
        With pvt
            'Get update status and suspend updates
            b_mu = .ManualUpdate
            .ManualUpdate = True

            .RowAxisLayout xlTabularRow
            .ColumnGrand = True
            .RowGrand = True
            .HasAutoFormat = False
            .ShowDrillIndicators = False

            For Each pvtField In pvt.PivotFields
                If pvtField.Orientation = xlRowField Then
                    pvtField.RepeatLabels = True
                    pvtField.Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
                End If
            Next pvtField

            'Restore update status
            .ManualUpdate = b_mu

        End With
    End If

    StandardExit
    
End Sub
