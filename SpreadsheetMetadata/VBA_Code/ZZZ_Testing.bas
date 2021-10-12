Attribute VB_Name = "ZZZ_Testing"
Option Explicit


Sub CreateLambdaQueryStorage()

    Dim ls As ListStorage
    
    Set ls = New ListStorage
    
    ls.CreateStorageFromPowerQuery ActiveWorkbook, "Lambdas", "Lambdas"

End Sub
