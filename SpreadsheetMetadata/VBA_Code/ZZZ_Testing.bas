Attribute VB_Name = "ZZZ_Testing"
'Copied from ExcelVbaCodeLibrary on <enter date here>

Option Explicit

Sub TestVariantArray()

    Dim wkb As Workbook
    Dim rng As Range
    Dim v As Variant
    Dim colProperty As Variant
    Dim colValue As Variant
    Dim unq As Variant
    Dim rv As Variant
    
    Set wkb = Application.Workbooks("PivotReportExample.xlsm")
    Set rng = wkb.Sheets("ReportPRoperties").ListObjects("tbl_ReportProperties").DataBodyRange
    v = rng
    colProperty = rng.Columns(4)
    colValue = rng.Columns(5)
    
    rv = WorksheetFunction.Xlookup("Category", colProperty, colValue)

End Sub



Sub TestListModules()

Dim vbcomp As VBComponent

For Each vbcomp In ThisWorkbook.VBProject.VBComponents

    'if normal Module
    'If vbcomp.Type = vbext_ct_StdModule Then
    If vbcomp.Type = vbext_ct_ClassModule Then

        Debug.Print vbcomp.Name
        
    End If
Next vbcomp

End Sub


Sub TestInsertProcedureCode(ByVal wb As Workbook, ByVal InsertToModuleName As String)

    Dim VBCM As CodeModule
    Dim InsertLineIndex As Long

    'On Error Resume Next
    Set VBCM = wb.VBProject.VBComponents(InsertToModuleName).CodeModule
    If Not VBCM Is Nothing Then
        With VBCM
            
            .InsertLines 1, "'Copied from ExcelVbaCodeLibrary on <enter date here>"
            .InsertLines 2, ""
        
'            InsertLineIndex = .CountOfLines + 1
'            ' customize the next lines depending on the code you want to insert
'            .InsertLines InsertLineIndex, "Sub NewSubName()" & Chr(13)
'            InsertLineIndex = InsertLineIndex + 1
'            .InsertLines InsertLineIndex, _
'                "    Msgbox ""Hello World!"",vbInformation,""Message Box Title""" & Chr(13)
'            InsertLineIndex = InsertLineIndex + 1
'            .InsertLines InsertLineIndex, "End Sub" & Chr(13)
            ' no need for more customizing
        End With
        Set VBCM = Nothing
    End If
    
    On Error GoTo 0

End Sub

Sub CallTestInsertProcudureCode()
    TestInsertProcedureCode ActiveWorkbook, "ZZZ_Testing"
End Sub


Sub TestDownloadFile()

    Dim myURL As String
    Dim oStream
    myURL = "https://raw.githubusercontent.com/charl-potgieter/SpreadsheetBI/master/SpreadsheetMetadata/VBA_Code/EntryPoints.bas"
    
    Dim WinHttpReq As Object
    Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")
    WinHttpReq.Open "GET", myURL, False
    WinHttpReq.send
    
    If WinHttpReq.Status = 200 Then
        Set oStream = CreateObject("ADODB.Stream")
        oStream.Open
        oStream.Type = 1
        oStream.Write WinHttpReq.responseBody
        oStream.SaveToFile "C:\Users\charl\Dropbox\file.bas", 2 ' 1 = no overwrite, 2 = overwrite
        oStream.Close
    End If

End Sub



Sub TestConditionalFormatting()

    Dim rng As Range
    Dim i As Integer
    Dim ErrorFormatCondition As FormatCondition
    Dim ErrorCheckFormatConditionIsError As FormatCondition
    Dim ErrorCheckFormatConditionNotSet As FormatCondition
    
    Set rng = ActiveCell
    
    'Delete existing conditional formatting - needs to be done
    'in reverse order to prevent errors
    For i = rng.FormatConditions.Count To 1 Step -1
        rng.FormatConditions(i).Delete
    Next i
    
    Set ErrorCheckFormatConditionNotSet = rng.FormatConditions.Add( _
        Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Error checking not set""")
    With ErrorCheckFormatConditionNotSet.Font
        .Bold = True
        .Italic = False
        .Color = RGB(255, 122, 0)
        .TintAndShade = 0
    End With

    Set ErrorCheckFormatConditionIsError = rng.FormatConditions.Add( _
        Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Workbook error - see index tab""")
    With ErrorCheckFormatConditionIsError.Font
        .Bold = True
        .Italic = False
        .Color = RGB(255, 0, 0)
        .TintAndShade = 0
    End With


End Sub
