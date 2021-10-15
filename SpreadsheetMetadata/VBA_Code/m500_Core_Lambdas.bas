Attribute VB_Name = "m500_Core_Lambdas"
Option Explicit
Option Private Module

Function GetLambdaFromUser(ByVal wkb As Workbook) As LambdaFormulaDetails

    Dim uf As uf_LambdaFunctionWizard
    Dim LambdaCategories As Variant
    Dim Lambdas As Dictionary
    Dim Storage
    Dim YesNoResponse As Integer
    
    Set Storage = AssignLambdaStorage
    Set uf = New uf_LambdaFunctionWizard
    
    Set Lambdas = ReadLambdaFormulaDetails(Storage)
    Set uf.LambdaDetails = Lambdas
    
    ReadUniqueLambdaCategories Storage, LambdaCategories
    uf.Categories = LambdaCategories
    
    uf.Show
    
    If uf.UserSelectedCancel Then
        Set GetLambdaFromUser = Nothing
        GoTo Exitpoint
    Else
        Set GetLambdaFromUser = uf.SelectedLambdaDetails
    End If

    
Exitpoint:
    On Error Resume Next
    Unload uf
    On Error GoTo 0
    
    Set uf = Nothing
    Set Storage = Nothing
    Set Lambdas = Nothing

End Function


Function GetLambdaParametersFromUser(ByVal SelectedLambda As LambdaFormulaDetails) As Variant

    Dim uf As uf_LambdaParameters
    
    Set uf = New uf_LambdaParameters
    uf.LambdaName = SelectedLambda.Name
    Set uf.ParameterDescriptions = SelectedLambda.ParameterDescriptions
    uf.Show
    
    If Not uf.UserSelectedCancel Then
        GetLambdaParametersFromUser = uf.OrderedParameterValues
    Else
        GetLambdaParametersFromUser = Empty
    End If
    
Exitpoint:
    On Error Resume Next
    Unload uf
    On Error GoTo 0
    
    Set uf = Nothing

End Function


Sub AddLambdaToWorkbook(ByVal wkb As Workbook, ByVal Lambda As LambdaFormulaDetails)

    Dim nm As Name
    Dim YesNoResponse As Integer

    Select Case True
        
        Case Not WorkbookLevelRangeNameExists(wkb, Lambda.Name)
            Set nm = wkb.Names.Add( _
                Name:=Lambda.Name, _
                RefersTo:=Lambda.RefersTo)
            nm.Comment = Lambda.Description
        
        Case CleanTrim(wkb.Names(Lambda.Name).RefersTo) <> CleanTrim(Lambda.RefersTo)
            YesNoResponse = MsgBox( _
                Prompt:="The stored lambda definition does not match that per workbook. " & vbCrLf & _
                    "Update definition in workbook?", _
                Buttons:=vbYesNo)
            If YesNoResponse = vbYes Then
                Set nm = wkb.Names(Lambda.Name)
                nm.RefersTo = Lambda.RefersTo
                nm.Comment = Lambda.Description
            End If
    
    End Select
End Sub


Sub WriteLambdaToCell(ByVal Lambda As LambdaFormulaDetails, ByVal rng As Range, _
    ByVal Params As Variant)

    Dim FormulaStr As String
    Dim i As Integer
    Dim Key As Variant
    Dim ParameterValue As String
    
    FormulaStr = "=" & Lambda.Name & "("
    For i = LBound(Params) To UBound(Params)
        If i <> LBound(Params) Then
            FormulaStr = FormulaStr & ","
        End If
        If Params(i) <> "" Then
            ParameterValue = Params(i)
            If Not StringIsARangeReference(ParameterValue) Then
                ParameterValue = """" & ParameterValue & """"
            End If
            FormulaStr = FormulaStr & ParameterValue
        End If
    Next i
    FormulaStr = FormulaStr & ")"
    rng.Formula2 = FormulaStr

End Sub
