VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListStorage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Use "This" declaration as an easy way to get intellisense to the classes private variables
'https://rubberduckvba.wordpress.com/2020/02/27/vba-classes-gateway-to-solid/
Private Type ListObjectStorage
    Sheet As Worksheet
    ListObj As ListObject
End Type
Private This As ListObjectStorage

Public Function CreateStorage(ByRef wkb As Workbook, ByVal sName As String, Headings() As String) As Boolean
'Does nothing and returns FALSE if sheet or listobject already exist
'Creates a new sheet named "sName" and listobject named "tbl_sName" in wkb and returns TRUE
    
    If StorageAlreadyExists(wkb, sName) Then
        CreateStorage = False
    Else
        CreateStorage = True
        Set This.Sheet = CreateListStorageSheet(wkb, sName)
        FormatListStorageSheet This.Sheet
        Set This.ListObj = CreateListStorageListObject(This.Sheet, Headings())
        FormatStorageListObject This.ListObj
    End If

End Function

Public Property Get Name() As String
    Name = This.Sheet.Name
End Property



Private Function StorageAlreadyExists(ByRef wkb As Workbook, ByVal sName As String) As Boolean

    Dim sht As Worksheet
    Dim sTestSheet As String
    Dim sTestListObject As String

    'Check for existing sheet
    On Error Resume Next
    sTestSheet = wkb.Sheets(sName).Name
    If Err.Number = 0 Then
        StorageAlreadyExists = True
    End If
    
    'Check for existing listobject
    For Each sht In wkb.Worksheets
        On Error Resume Next
        sTestListObject = sht.ListObjects("tbl_" & sName)
        If Err.Number = 0 Then
            StorageAlreadyExists = True
        End If
        On Error GoTo 0
    Next sht
        

End Function



Private Function CreateListStorageSheet(ByRef wkb As Workbook, ByVal sName As String) As Worksheet

    Dim sht As Worksheet
    Set sht = wkb.Sheets.Add(After:=wkb.Sheets(wkb.Sheets.Count))
    sht.Name = sName
    Set CreateListStorageSheet = sht
    
End Function

Private Sub FormatListStorageSheet(ByRef sht As Worksheet)
'Applies my preferred sheet formattting

    sht.Activate
    
    sht.Cells.Font.Name = "Calibri"
    sht.Cells.Font.Size = 11
      
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    sht.DisplayPageBreaks = False
    sht.Columns("A:A").ColumnWidth = 4
    
    sht.Names.Add Name:="SheetHeading", RefersTo:="=$B$2"
    sht.Names.Add Name:="SheetCategory", RefersTo:="=$A$1"
    
    With sht.Range("SheetHeading")
        .Value = sht.Name
        .Font.Bold = True
        .Font.Size = 16
    End With
    
    With sht.Range("SheetCategory")
        .Value = "List Storage"
        .Font.Color = RGB(170, 170, 170)
        .Font.Size = 8
    End With
    

End Sub

Private Function CreateListStorageListObject(ByRef sht As Worksheet, Headings() As String) As ListObject

    Dim i As Integer
    Const iTableRowStart As Integer = 7
    Const iTableColStart As Integer = 2
    
    For i = LBound(Headings) To UBound(Headings)
        sht.Cells(iTableRowStart, iTableColStart + i).Value = Headings(i)
    Next i

    Set CreateListStorageListObject = sht.ListObjects.Add(xlSrcRange, _
        sht.Cells(iTableRowStart, iTableColStart).CurrentRegion, , xlYes)

End Function


Private Sub FormatStorageListObject(ByRef lo As ListObject)

    Dim sty As TableStyle
    Dim wkb As Workbook
    
    Set wkb = lo.Parent.Parent
    
    On Error Resume Next
    wkb.TableStyles.Add ("SpreadsheetBiStyle")
    On Error GoTo 0
    Set sty = wkb.TableStyles("SpreadsheetBiStyle")
    
    'Set Header Format
    With sty.TableStyleElements(xlHeaderRow)
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Borders.item(xlEdgeTop).LineStyle = xlSolid
        .Borders.item(xlEdgeTop).Weight = xlMedium
        .Borders.item(xlEdgeBottom).LineStyle = xlSolid
        .Borders.item(xlEdgeBottom).Weight = xlMedium
    End With

    'Set row stripe format
    sty.TableStyleElements(xlRowStripe1).Interior.Color = RGB(217, 217, 217)
    sty.TableStyleElements(xlRowStripe2).Interior.Color = RGB(255, 255, 255)
    
    'Set whole table bottom edge format
    sty.TableStyleElements(xlWholeTable).Borders.item(xlEdgeBottom).LineStyle = xlSolid
    sty.TableStyleElements(xlWholeTable).Borders.item(xlEdgeBottom).Weight = xlMedium

    
    'Apply custom style and set other attributes
    lo.TableStyle = "SpreadsheetBiStyle"
    With lo.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
    End With
    
    If Not lo.DataBodyRange Is Nothing Then
        lo.DataBodyRange.EntireColumn.AutoFit
    End If
    

End Sub
