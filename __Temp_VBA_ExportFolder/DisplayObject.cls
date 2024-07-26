VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DisplayObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder("Formatting")
Option Explicit

Private Type TypeDisplayObject
    InputRange As Range
End Type
Private this As TypeDisplayObject



Public Sub CreateFromRange(ByVal InputRange As Range)
    Set this.InputRange = InputRange
End Sub


Public Sub SetNumberFormat(ByVal CustomFormat As String)

    ' TODO - need different behaviour for listobjects columns and pivotfields
    this.InputRange.NumberFormat = CustomFormat

End Sub



Private Function CellsAreContainedInSingleListObject() As Boolean

    Dim CurrentListObject As ListObject
    Dim NumberOfCellsInRange As Long
    Dim NumberOfCellsInRangeInsideListObect As Long

    Set CurrentListObject = ListObjectInFirstCellOfRange(this.InputRange)
    If CurrentListObject Is Nothing Then
        CellsAreContainedInSingleListObject = False
    Else
        NumberOfCellsInRange = this.InputRange.Cells.Count
        NumberOfCellsInRangeInsideListObect = Intersect(CurrentListObject.Range, this.InputRange).Cells.Count
        CellsAreContainedInSingleListObject = (NumberOfCellsInRange = NumberOfCellsInRangeInsideListObect)
    End If
        
End Function



Private Function ListObjectInFirstCellOfRange(ByVal InputRange As Range) As ListObject
    Dim FirstCell As Range
    Set FirstCell = InputRange.Cells(1, 1)
    Set ListObjectInFirstCellOfRange = FirstCell.ListObject
End Function




Private Function EachCellIsInPivotFieldDataRange(ByVal TargetRange As Range) As Boolean

'   TODO  - this sub
'    Dim pvt As PivotTable
'    Dim TargetCell As Range
'    Dim ReturnValue As Boolean
'
'    ReturnValue = True
'    On Error Resume Next
'
'    For Each TargetCell In TargetRange
'        Set pvt = ActiveCell.PivotTable
'        ReturnValue = ReturnValue And (Len(pvt.Name) <> 0)
'    Next TargetCell
'
'    On Error GoTo 0
'    EachCellIsInPivotFieldDataRange = ReturnValue
    
    EachCellIsInPivotFieldDataRange = False
    
End Function


Private Function EachCellIsInListObjectDataRange() As Boolean

    Dim FirstCellInSelection As Range
    Dim SelectedListObject As ListObject
    Dim ReturnValue As Boolean
    Dim NumberOfCellsInTargetRange As Long
    Dim NumberOfCellsInTargetIncludedInListObect As Long
    

    Set FirstCellInSelection = this.InputRange.Cells(1, 1)
    Set SelectedListObject = FirstCellInSelection.ListObject
    
    If SelectedListObject Is Nothing Then
        ReturnValue = False
    Else
        NumberOfCellsInTargetRange = this.InputRange.Cells.Count
        NumberOfCellsInTargetIncludedInListObect = Intersect(SelectedListObject.DataBodyRange, this.InputRange).Cells.Count
        ReturnValue = (NumberOfCellsInTargetRange = NumberOfCellsInTargetIncludedInListObect)
    End If
        
   EachCellIsInListObjectDataRange = ReturnValue
        
End Function

