Attribute VB_Name = "m020_DATA_ACCESS_Lambdas"
Option Explicit
Option Private Module


Function AssignLambdaStorage()

    Dim LambdaStorage As ListStorage

    Set LambdaStorage = New ListStorage
    LambdaStorage.AssignStorage ThisWorkbook, "Lambdas"
    Set AssignLambdaStorage = LambdaStorage

End Function




Sub ReadUniqueLambdaCategories(ByVal LambdaStorage, ByRef LambdaCategories As Variant)

    Dim Storage As ListStorage

    Set Storage = LambdaStorage
    LambdaCategories = Storage.ItemsInField(sFieldName:="Category", bUnique:=True, bSorted:=True)

End Sub



Function ReadLambdaFormulaDetails(ByVal LambdaStorage) As Dictionary

    Dim LambdaFormulas As Dictionary
    Dim i As Integer
    Dim j As Integer
    Dim Storage As ListStorage
    Dim NumberOfLambdas As Integer
    Dim LambdaFormulaDetails As LambdaFormulaDetails
    Dim ParameterAndDescriptions() As String
    Dim ParamaterDescriptionPairString As String
    Dim NumberOfParameters As Integer
    Dim ParameterName As String
    Dim ParameterDescription As String

    'Below is performed to enable intellisense that is not available for variant type parameter
    Set Storage = LambdaStorage

    NumberOfLambdas = Storage.NumberOfRecords
    Set LambdaFormulas = New Dictionary

    For i = 0 To NumberOfLambdas - 1
        Set LambdaFormulaDetails = New LambdaFormulaDetails
        With LambdaFormulaDetails
            .Name = Storage.FieldItemByIndex("Name", i + 1)
            .RefersTo = Storage.FieldItemByIndex("RefersTo", i + 1)
            .Category = Storage.FieldItemByIndex("Category", i + 1)
            .Author = Storage.FieldItemByIndex("Author", i + 1)
            .Description = Storage.FieldItemByIndex("Description", i + 1)
            .GitUrl = Storage.FieldItemByIndex("GitUrl", i + 1)

            ParamaterDescriptionPairString = Storage.FieldItemByIndex("PipeDelimitedParameterAndDescription", i + 1)
            NumberOfParameters = (Len(ParamaterDescriptionPairString) - _
                Len(Replace(ParamaterDescriptionPairString, "|", "")) + 1) / 2
            ReDim ParameterAndDescriptions(0 To NumberOfParameters * 2 - 1)
            ParameterAndDescriptions = Split(Storage.FieldItemByIndex("PipeDelimitedParameterAndDescription", i + 1), "|")
            Set .ParameterDescriptions = New Dictionary
            For j = 0 To NumberOfParameters - 1
                ParameterName = ParameterAndDescriptions(j * 2)
                ParameterDescription = ParameterAndDescriptions((j * 2) + 1)
                .ParameterDescriptions.Add ParameterName, ParameterDescription
            Next j
        End With
        LambdaFormulas.Add Storage.FieldItemByIndex("Name", i + 1), LambdaFormulaDetails
        Set LambdaFormulaDetails = Nothing
    Next i

    Set ReadLambdaFormulaDetails = LambdaFormulas


End Function


Sub RefreshLambdaLibrariesFromGithub()

    Dim ls As ListStorage
    
    Set ls = New ListStorage
    ls.AssignStorage ThisWorkbook, "Lambdas"
    ls.ListObj.Refresh

End Sub


