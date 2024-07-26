Attribute VB_Name = "m000_ENTRY_POINTS_Lambdas"
'@Folder "SpreadsheetBI"
Option Explicit
Option Private Module

Public Sub RefreshLambdaLibraries()

    StandardEntry
    RefreshLambdaLibrariesFromGithub
    StandardExit

End Sub


Public Sub InsertLambda()
Attribute InsertLambda.VB_ProcData.VB_Invoke_Func = "F\n14"

    Dim SelectedLambda As LambdaFormulaDetails
    Dim LambdaParameters As Variant
    Dim wkb As Workbook
    
    'StandardEntry
    Set wkb = ActiveWorkbook
    
    
    Set SelectedLambda = GetLambdaFromUser(wkb)
    
    If Not SelectedLambda Is Nothing Then
    
        'Turn screenupdating on to view "marching ants" around selected range
        Application.ScreenUpdating = True
        LambdaParameters = GetLambdaParametersFromUser(SelectedLambda)
        Application.ScreenUpdating = False
        
        If Not IsEmpty(LambdaParameters) Then
            AddLambdaToWorkbook wkb, SelectedLambda
            WriteLambdaToCell SelectedLambda, ActiveCell, LambdaParameters
        End If
        
    End If
    
Exitpoint:
    Set SelectedLambda = Nothing
    Set wkb = Nothing
    StandardExit

End Sub
