Attribute VB_Name = "ZZZ_Testing"
Option Explicit


Sub CreateLambdaQueryStorage()

    Dim ls As ListStorage
    
    Set ls = New ListStorage
    
    ls.CreateStorageFromPowerQuery ActiveWorkbook, "Lambdas", "Lambdas"

End Sub



Sub TEst()

    Dim Response
    
    Response = Application.InputBox(Prompt:="test", Type:=0)

End Sub



Sub test2()

    Dim uf As uf_LambdaParameters
    
    Set uf = New uf_LambdaParameters
    
    uf.Show

End Sub



Sub testlist()
    
    Dim VBProj As VBIDE.VBProject
    Dim VBComp As VBIDE.VBComponent
    
    'On Error Resume Next
    Set VBProj = ThisWorkbook.VBProject
    
    For Each VBComp In VBProj.VBComponents
        Debug.Print VBComp.Name & "|"; VBComp.Type
    Next VBComp
    
    
'    Set VBComp = VBProj.VBComponents(sModuleName)
'    VBProj.VBComponents.Remove VBComp
'    DeleteModule = (Err.Number = 0)
'    On Error GoTo 0
    
End Sub
