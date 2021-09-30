Attribute VB_Name = "m020_DATA_ACCESS_Lambdas"
Option Explicit
Option Private Module

Private Const csRepoGitRepoStorageName = "__GitRepos"
Private Const csLambdaStorageName = "__Lambdas"


Function AssignGitRepoStorage()

    Dim GitRepoStorage As ListStorage
    Dim wkb As Workbook
    
    Set wkb = ActiveWorkbook
    Set GitRepoStorage = New ListStorage
    
    
    If Not (GitRepoStorage.StorageAlreadyExists(wkb, csRepoGitRepoStorageName)) Then
        GitRepoStorage.CreateStorage wkb, csRepoGitRepoStorageName, Array("RepoUrl")
    Else
        GitRepoStorage.AssignStorage wkb, csRepoGitRepoStorageName
    End If

    Set AssignGitRepoStorage = GitRepoStorage

End Function


Function AssignLambdaStorage()

    Dim LambdaStorage As ListStorage
    Dim wkb As Workbook
    
    Set wkb = ActiveWorkbook
    Set LambdaStorage = New ListStorage
    
    
    If Not (LambdaStorage.StorageAlreadyExists(wkb, csLambdaStorageName)) Then
        LambdaStorage.CreateStorage wkb, csLambdaStorageName, Array("Name", "RefersTo", "Category", "Author", "Description", "ParameterDescription", "URL")
    Else
        LambdaStorage.AssignStorage wkb, csLambdaStorageName
    End If

    Set AssignLambdaStorage = LambdaStorage

End Function




Function GitRepoStorageExists() As Boolean

    Dim GitRepoStorage As ListStorage
    Dim wkb As Workbook
    
    Set wkb = ActiveWorkbook
    Set GitRepoStorage = New ListStorage
    
    GitRepoStorageExists = (GitRepoStorage.StorageAlreadyExists(wkb, csRepoGitRepoStorageName))

End Function


Function RepoAlreadyExistsInStorage(ByVal sRepoUrl As String, ByVal GitRepoStorage) As Boolean

    Dim ArrayOfReposAlreadyAdded
    Dim i As Integer

    RepoAlreadyExistsInStorage = False

    If Not GitRepoStorage.IsEmpty Then
        ArrayOfReposAlreadyAdded = GitRepoStorage.ItemsInField("RepoUrl")
        i = LBound(ArrayOfReposAlreadyAdded)
        Do While i <= UBound(ArrayOfReposAlreadyAdded) And Not RepoAlreadyExistsInStorage
            RepoAlreadyExistsInStorage = (UCase(sRepoUrl) = UCase(ArrayOfReposAlreadyAdded(i)))
            i = i + 1
        Loop
    End If

End Function


Sub AddRepoToStorage(ByVal sRepoUrl As String, ByRef GitRepoStorage)

    Dim RepoUrlDictionary As Dictionary
    
    Set RepoUrlDictionary = New Dictionary
    RepoUrlDictionary.Add key:="RepoURL", Item:=sRepoUrl
    GitRepoStorage.InsertFromDictionary RepoUrlDictionary

End Sub


Sub ReadRepoList(ByRef sRepoList() As String, ByVal GitRepoStorage)

    Dim NumberOfRepos As Integer
    Dim Storage As ListStorage
    Dim i As Integer
    
    'Below is performed to enable intellisense that is not available for variant tpye parameter
    Set Storage = GitRepoStorage
    
    NumberOfRepos = Storage.NumberOfRecords
    
    ReDim sRepoList(0 To NumberOfRepos - 1)

    For i = 1 To NumberOfRepos
        sRepoList(i - 1) = Storage.FieldItemByIndex("RepoUrl", i)
    Next i

End Sub


Sub ImportDataIntoLambdaStorage(ByRef sRepoList() As String, ByVal LambdaStorage, _
    ByVal LambdaXmlMap As XmlMap)

    Dim sRepoUrl As String
    Dim Storage As ListStorage
    Dim wkb As Workbook
    Dim i As Integer

    'Below is performed to enable intellisense that is not available for variant type parameter
    Set Storage = LambdaStorage
    
    Set wkb = Storage.ListObj.Parent.Parent
    
    'Assign XML map to storage list object
    With Storage.ListObj
        .ListColumns("Name").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Name"
        .ListColumns("RefersTo").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/RefersTo"
        .ListColumns("Category").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Category"
        .ListColumns("Author").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Author"
        .ListColumns("Description").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Description"
        .ListColumns("ParameterDescription").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/ParameterDescription"
    End With


    Storage.ClearData
    For i = LBound(sRepoList) To UBound(sRepoList)
        sRepoUrl = sRepoList(i)
        wkb.XmlMaps(gcsLambdaXmlMapName).Import URL:=sRepoUrl, Overwrite:=False
        If LambdaStorage.ListObj.ListColumns("URL").DataBodyRange.Cells.Count = 1 Then
            'for some reason SpecialCells does not seem to work for one cell
            LambdaStorage.ListObj.ListColumns("URL").DataBodyRange = sRepoUrl
        Else
            LambdaStorage.ListObj.ListColumns("URL").DataBodyRange.SpecialCells(xlCellTypeBlanks) = sRepoUrl
        End If
    Next i

    With Storage.ListObj
        .ListColumns("Name").Range.ColumnWidth = 25
        .ListColumns("RefersTo").Range.ColumnWidth = 90
        .ListColumns("Category").Range.ColumnWidth = 25
        .ListColumns("Author").Range.ColumnWidth = 25
        .ListColumns("Description").Range.ColumnWidth = 40
        .ListColumns("URL").Range.ColumnWidth = 90
        .DataBodyRange.HorizontalAlignment = xlLeft
        .DataBodyRange.VerticalAlignment = xlTop
        .DataBodyRange.WrapText = True
        .DataBodyRange.EntireRow.AutoFit
    End With
    

End Sub



Sub ReadLambdaFormulaDetails(ByVal LambdaStorage, ByRef LambdaFormulas As Dictionary)

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
            .RefersTo = Storage.FieldItemByIndex("RefersTo", i + 1)
            .Category = Storage.FieldItemByIndex("Category", i + 1)
            .Author = Storage.FieldItemByIndex("Author", i + 1)
            .Description = Storage.FieldItemByIndex("Description", i + 1)
            .URL = Storage.FieldItemByIndex("Name", i + 1)
        
            ParamaterDescriptionPairString = Storage.FieldItemByIndex("ParameterDescription", i + 1)
            NumberOfParameters = (Len(ParamaterDescriptionPairString) - _
                Len(Replace(ParamaterDescriptionPairString, "|", "")) + 1) / 2
            ReDim ParameterAndDescriptions(0 To NumberOfParameters * 2 - 1)
            ParameterAndDescriptions = Split(Storage.FieldItemByIndex("ParameterDescription", i + 1), "|")
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

End Sub



Sub ReadLambdaNamesPerCategory(ByVal LambdaStorage, ByRef LambdaNames, ByVal Category As String)

    Dim Storage As ListStorage
    Dim sFilterString As String
    
    Set Storage = LambdaStorage
    If Category = "All" Then
        LambdaNames = Storage.ItemsInField("Name")
    Else
        sFilterString = "[Category] = """ & Category & """"
        Storage.Filter sFilterString
        LambdaNames = Storage.ItemsInField(sFieldName:="Name", bFiltered:=True)
    End If
    
End Sub


Sub ReadUniqueLambdaCategories(ByVal LambdaStorage, ByRef LambdaCategories)

    Dim Storage As ListStorage
    
    Set Storage = LambdaStorage
    LambdaCategories = Storage.ItemsInField(sFieldName:="Category", bUnique:=True)
    
End Sub



