Attribute VB_Name = "m500_Core_Lambdas"
Option Explicit
Option Private Module

Sub SetupGeneratorCategorySheet(ByRef sht As Worksheet, ByRef lo As ListObject)

    sht.Name = "Categories"
    sht.Range("B5") = "Categories"
    
    Set lo = sht.ListObjects.Add(xlSrcRange, _
        sht.Range("B5").CurrentRegion, , xlYes)
    lo.Name = "tbl_Categories"
    lo.ListColumns("Categories").Range.ColumnWidth = 50
    
    'Force existence of DataBodyRange
    lo.HeaderRowRange.Cells(1).Offset(1, 0).Value = " "
    
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



Sub SetupGeneratorLambdaSheet(ByRef sht As Worksheet, ByRef lo As ListObject)

    Dim wkb As Workbook

    sht.Name = "Lambdas"
    sht.Range("B5") = "Name"
    sht.Range("C5") = "RefersTo"
    sht.Range("D5") = "Category"
    sht.Range("E5") = "Author"
    sht.Range("F5") = "Description"
    sht.Range("G5") = "ParameterDescription"
    
    Set lo = sht.ListObjects.Add(xlSrcRange, _
        sht.Range("B5").CurrentRegion, , xlYes)
    lo.Name = "tbl_Lambdas"
    
    'Force existence of DataBodyRange
    lo.HeaderRowRange.Cells(1).Offset(1, 0).Value = " "
    
    With lo
        .ListColumns("Name").Range.ColumnWidth = 25
        .ListColumns("RefersTo").Range.ColumnWidth = 90
        .ListColumns("Category").Range.ColumnWidth = 25
        .ListColumns("Author").Range.ColumnWidth = 25
        .ListColumns("Description").Range.ColumnWidth = 40
        .ListColumns("ParameterDescription").Range.ColumnWidth = 70
        .DataBodyRange.HorizontalAlignment = xlLeft
        .DataBodyRange.VerticalAlignment = xlTop
        .DataBodyRange.WrapText = True
        .DataBodyRange.EntireRow.AutoFit
    End With
    
    
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
    
    'Add Comment re Category data validation
    With lo.ListColumns("Category").Range.Cells(1)
        .AddComment
        .Comment.Visible = True
        .Comment.Text Text:= _
            "Drop down data validation is based on categories as captured " & _
            "in the second tab of this workbook."
        .Comment.Shape.Left = 500
        .Comment.Shape.Top = 10
        .Comment.Shape.Width = 200
        .Comment.Shape.Height = 40
    End With
    
    'Add Comment re ParameterDescription
    With lo.ListColumns("ParameterDescription").Range.Cells(1)
        .AddComment
        .Comment.Visible = True
        .Comment.Text Text:= _
            "Enter as a pipe delimited string of name description pairs.  e.g." & _
            "ParamaterName|Description|ParamaterName|Description"
        .Comment.Shape.Left = 1200
        .Comment.Shape.Top = 10
        .Comment.Shape.Width = 300
        .Comment.Shape.Height = 40
    End With
    
    
    'Add validation to categories field on LambaStorage
    Set wkb = sht.Parent
    
    wkb.Names.Add Name:="Val_Categories", RefersToR1C1:="=tbl_Categories[Categories]"
    lo.ListColumns("Category").DataBodyRange.Validation.Add _
        Type:=xlValidateList, Formula1:="=Val_Categories", AlertStyle:=xlValidAlertStop
        
    
    
End Sub



Sub FormatGeneratorListObject(ByRef lo As ListObject)

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
    lo.TableStyle = "SpreadsheetBiStyle"
    With lo.HeaderRowRange
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
        .Orientation = 0
    End With

End Sub



Function CreateLambdaXmlMap(ByVal wkb As Workbook) As XmlMap

    Dim sMap As String
    Dim LambdaXmlMap As XmlMap


    On Error Resume Next
    wkb.XmlMaps(gcsLambdaXmlMapName).Delete
    On Error GoTo 0

    'Excel needs two elements in map such a below in order to work out the schema
    sMap = "<LambdaDocument> " & vbCrLf & _
            " <Record> " & vbCrLf & _
            "    <Name></Name><RefersTo></RefersTo><Category></Category><Author></Author><Description></Description><ParameterDescription></ParameterDescription> " & vbCrLf & _
            " </Record> " & vbCrLf & _
            " <Record> " & vbCrLf & _
            "    <Name></Name><RefersTo></RefersTo><Category></Category><Author></Author><Description></Description><ParameterDescription></ParameterDescription> " & vbCrLf & _
            " </Record> " & vbCrLf & _
            "</LambdaDocument>"

    'Create XML map in sht parent
    On Error Resume Next
    wkb.XmlMaps(gcsLambdaXmlMapName).Delete
    On Error GoTo 0
    Set LambdaXmlMap = wkb.XmlMaps.Add(sMap, "LambdaDocument")
    LambdaXmlMap.Name = gcsLambdaXmlMapName

    Set CreateLambdaXmlMap = LambdaXmlMap

End Function





Function WorkbookIsValidForLambdaXmlExport(ByVal wkb As Workbook) As Boolean

    WorkbookIsValidForLambdaXmlExport = True
    
    On Error Resume Next
    If Err.Number <> 0 Then
        MsgBox ("This workbook is not in the correct format to export lambda functions")
        WorkbookIsValidForLambdaXmlExport = False
    End If
    On Error GoTo 0
    
    If wkb.Path = "" Then
        MsgBox ("Workbook needs to be saved before generation of output")
        WorkbookIsValidForLambdaXmlExport = False
    End If
    

End Function


Sub WriteHumanReadableLambdaInventory(ByRef loLambdas As ListObject, _
    ByVal sFilePath As String)
'Requires reference to Microsoft Scripting Runtime
'Writes sStr to a text file
'*** THIS WILL OVERWRITE ANY CURRENT CONTENT OF THE FILE ***

    Dim fso As Object
    Dim oFile As Object
    Dim i As Long

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set oFile = fso.CreateTextFile(sFilePath)

    For i = 1 To loLambdas.DataBodyRange.Rows.Count
        oFile.WriteLine ("/*------------------------------------------------------------------------------------------------------------------")
        oFile.WriteLine ("      Formula Name:   " & loLambdas.ListColumns("Name").DataBodyRange.Cells(i))
        oFile.WriteLine ("      Category:       " & loLambdas.ListColumns("Category").DataBodyRange.Cells(i))
        oFile.WriteLine ("      Autohor:        " & loLambdas.ListColumns("Author").DataBodyRange.Cells(i))
        oFile.WriteLine ("      Description:        " & loLambdas.ListColumns("Description").DataBodyRange.Cells(i))
        oFile.WriteLine ("------------------------------------------------------------------------------------------------------------------/*")
        oFile.WriteLine (loLambdas.ListColumns("RefersTo").DataBodyRange.Cells(i))
        oFile.WriteLine (vbCrLf)
    Next i

    oFile.Close
    Set fso = Nothing
    Set oFile = Nothing

End Sub


Sub DeleteExistingLambdaFormulas(ByVal wkb As Workbook)

    Dim nm As Name

    For Each nm In ActiveWorkbook.Names
        If Left(nm.Comment, 20) = gcsCommentPrefix Then
            nm.Delete
        End If
    Next nm

End Sub



Sub CreateLambdaFormulas(ByVal wkb As Workbook, ByRef LambdaFormulas As Dictionary)

    Dim FormulaName
    Dim CleanedName As String
    Dim CleanedFormula As String
    Dim nm As Name

    For Each FormulaName In LambdaFormulas.Keys

        CleanedName = Replace( _
            WorksheetFunction.Clean(WorksheetFunction.Trim(FormulaName)), _
            " ", _
            "")
        CleanedFormula = WorksheetFunction.Clean(WorksheetFunction.Trim(LambdaFormulas(FormulaName).RefersTo))

        Set nm = wkb.Names.Add( _
            Name:=CleanedName, _
            RefersTo:=CleanedFormula)
        nm.Comment = gcsCommentPrefix

    Next FormulaName

End Sub





