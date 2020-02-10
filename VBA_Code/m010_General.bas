Attribute VB_Name = "m010_General"
Option Explicit
Option Private Module
Global Const gcsMenuName As String = "SpreadsheetBI"


Sub FormatSheet(ByRef sht As Worksheet)
'Applies my preferred sheet formattting

    sht.Activate
    
    sht.Range("A1").Font.Color = RGB(170, 170, 170)
    sht.Range("A1").Font.Size = 8
    
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 80
    sht.DisplayPageBreaks = False
    sht.Columns("A:A").ColumnWidth = 4
    
    If SheetLevelRangeNameExists(sht, "SheetHeading") Then
        sht.Names("SheetHeading").Delete
    End If
    sht.Names.Add Name:="SheetHeading", RefersTo:="=$B$2"
    
    With sht.Range("SheetHeading")
        If .value = "" Then
            .value = "Heading"
        End If
        .Font.Bold = True
        .Font.Size = 16
    End With

End Sub




Sub FormatTable(lo As ListObject)

    Dim sty As TableStyle
    
    On Error Resume Next
    ActiveWorkbook.TableStyles("CustomTableStyle").Delete
    On Error GoTo 0
    
    Set sty = ActiveWorkbook.TableStyles.Add("CustomTableStyle")
    
    'Set Header Format
    With sty.TableStyleElements(xlHeaderRow)
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Borders.Item(xlEdgeTop).LineStyle = xlSolid
        .Borders.Item(xlEdgeTop).Weight = xlMedium
        .Borders.Item(xlEdgeBottom).LineStyle = xlSolid
        .Borders.Item(xlEdgeBottom).Weight = xlMedium
    End With

    'Set row stripe format
    sty.TableStyleElements(xlRowStripe1).Interior.Color = RGB(217, 217, 217)
    sty.TableStyleElements(xlRowStripe2).Interior.Color = RGB(255, 255, 255)
    
    'Set whole table bottom edge format
    sty.TableStyleElements(xlWholeTable).Borders.Item(xlEdgeBottom).LineStyle = xlSolid
    sty.TableStyleElements(xlWholeTable).Borders.Item(xlEdgeBottom).Weight = xlMedium

    
    'Apply custom style and set other attributes
    lo.TableStyle = "CustomTableStyle"
    With lo.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
    End With
    
    lo.DataBodyRange.EntireColumn.AutoFit


End Sub



Sub SetNumberFormat(sNumberFormat)

    If ActiveCellIsInPivotTable Then
        ActiveCell.PivotField.NumberFormat = sNumberFormat
    Else
        Selection.NumberFormat = sNumberFormat
    End If

End Sub



Function LooperValue(ByVal sItem As String) As String
'Precondition: tbl_LoopController exists in active workbook and contans column Item and Value
'This sub returns Value for corresponding sItem

    Dim sFormulaString As String
    
    sFormulaString = "=INDEX(tbl_LoopController[Value], MATCH(""" & sItem & """, tbl_LoopController[Item], 0))"
    LooperValue = Application.Evaluate(sFormulaString)


End Function




Sub DeletePopUpMenu()
'Delete PopUp menu if it exists
    
    On Error Resume Next
    Application.CommandBars(gcsMenuName).Delete
    On Error GoTo 0
    
End Sub



Sub CreatePopUpMenu()

    Dim cb As CommandBar
    Dim MenuCategory As CommandBarPopup
    Dim MenuItem As CommandBarControl
    Dim i As Integer
    Dim sCurrentMenuCategory As String
    Dim sPreviousMenuCategory As String
    Dim lo As ListObject
    
    
    Set cb = Application.CommandBars.Add(Name:=gcsMenuName, Position:=msoBarPopup, _
                                     MenuBar:=False, Temporary:=True)
    Set lo = ThisWorkbook.Sheets("MenuGenerator").ListObjects("tbl_MenuGenerator")
    sPreviousMenuCategory = ""
    
    With cb
                                     
        For i = 1 To lo.DataBodyRange.Rows.Count
            sCurrentMenuCategory = lo.ListColumns("Category").DataBodyRange.Cells(i)
            If sCurrentMenuCategory <> sPreviousMenuCategory Then
                Set MenuCategory = .Controls.Add(Type:=msoControlPopup)
                MenuCategory.Caption = sCurrentMenuCategory
                sPreviousMenuCategory = sCurrentMenuCategory
            End If
        
            Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
            MenuItem.Caption = lo.ListColumns("Menu Item").DataBodyRange.Cells(i)
            MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & lo.ListColumns("Macro").DataBodyRange.Cells(i)
        
        Next i
                                     
                                     

'        'First add two buttons
'        With .Controls.Add(Type:=msoControlButton)
'            .Caption = "Button 1"
'            .FaceId = 71
'            .OnAction = "'" & ThisWorkbook.Name & "'!" & "TestMacro"
'        End With
'
'        With .Controls.Add(Type:=msoControlButton)
'            .Caption = "Button 2"
'            .FaceId = 72
'            .OnAction = "'" & ThisWorkbook.Name & "'!" & "TestMacro"
'        End With
'
'        'Second Add menu with two buttons
'        Set MenuItem = .Controls.Add(Type:=msoControlPopup)
'        With MenuItem
'            .Caption = "My Special Menu"
'
'            With .Controls.Add(Type:=msoControlButton)
'                .Caption = "Button 1 in menu"
'                .FaceId = 71
'                .OnAction = "'" & ThisWorkbook.Name & "'!" & "TestMacro"
'            End With
'
'            With .Controls.Add(Type:=msoControlButton)
'                .Caption = "Button 2 in menu"
'                .FaceId = 72
'                .OnAction = "'" & ThisWorkbook.Name & "'!" & "TestMacro"
'            End With
'        End With
'
'        'Third add one button
'        With .Controls.Add(Type:=msoControlButton)
'            .Caption = "Button 3"
'            .FaceId = 73
'            .OnAction = "'" & ThisWorkbook.Name & "'!" & "TestMacro"
'        End With

    End With
End Sub


Sub TestMacro()
    MsgBox "Hi There, greetings from the Netherlands"
End Sub




