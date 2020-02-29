Attribute VB_Name = "Module2"
Option Explicit

Sub Macro1()
Attribute Macro1.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro1 Macro
'

'
    Range(Selection, Selection.End(xlDown)).Select
    Range("D8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Range("D8").Select
    Range(Selection, Selection.End(xlDown)).Select
    Application.CutCopyMode = False
    Range("D8").Select
    ActiveCell.Formula2R1C1 = "=INDIRECT(""val_"" & RC[-1] & ""s"")"
    Range("D8").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=INDIRECT(""val_"" & C8 & ""s"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("D9").Select
End Sub
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=INDIRECT(""val_"" & C8 & ""s"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
End Sub
Sub Macro3()
Attribute Macro3.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro3 Macro
'

'
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=INDIRECT(""val_"" & C8 & ""s"")"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    Range("D8").Select
End Sub


Sub TestVal()

    Dim lo As ListObject
    Dim sRelativeReferenceOfDataFieldType As String
    Dim sValidationStr As String
    
    Set lo = ActiveWorkbook.Sheets("ReportFieldSettings").ListObjects(1)
    
    sRelativeReferenceOfDataFieldType = Replace(lo.ListColumns("Data Model Field Type").DataBodyRange.Cells(1).Address, "$", "")
    sValidationStr = "=INDIRECT(""val_"" & IF(" & sRelativeReferenceOfDataFieldType & " ="""", ""Measure"", " & sRelativeReferenceOfDataFieldType & ") & ""s"")"
    lo.ListColumns("Cube Field Name").DataBodyRange.Validation.Add _
        Type:=xlValidateList, Formula1:=sValidationStr
    

End Sub
