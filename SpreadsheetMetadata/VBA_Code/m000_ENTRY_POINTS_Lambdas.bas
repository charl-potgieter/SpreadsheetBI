Attribute VB_Name = "m000_ENTRY_POINTS_Lambdas"
Option Explicit

Public Const gcsLambdaXmlMapName As String = "LambdaMap"
Public Const gcsCommentPrefix = "<PowerFormulaImport>"

    

Public Sub CreateLambdaXmlGeneratorWorkbook()

    Dim shtCategories As Worksheet
    Dim shtLambdas As Worksheet
    Dim loCategories As ListObject
    Dim loLambdas As ListObject
    Dim wkb As Workbook
    Dim sht As Worksheet
    Dim LambdaXmlMap As XmlMap

    
    StandardEntry
    Set wkb = Workbooks.Add
    
    'Delete all sheets except 1
    For Each sht In wkb.Worksheets
        If sht.Index <> 1 Then sht.Delete
    Next sht
    
    Set shtCategories = wkb.Worksheets(1)
    SetupGeneratorCategorySheet shtCategories, loCategories
    FormatGeneratorListObject loCategories
    Set shtLambdas = wkb.Sheets.Add(Before:=wkb.Sheets(1))
    SetupGeneratorLambdaSheet shtLambdas, loLambdas
    FormatGeneratorListObject loLambdas
    
    Set LambdaXmlMap = CreateLambdaXmlMap(wkb)
    
    'Assign XML map to Lambdas list object
    With loLambdas
        .ListColumns("Name").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Name"
        .ListColumns("RefersTo").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/RefersTo"
        .ListColumns("Category").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Category"
        .ListColumns("Author").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Author"
        .ListColumns("Description").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/Description"
        .ListColumns("ParameterDescription").XPath.SetValue LambdaXmlMap, "/LambdaDocument/Record/ParameterDescription"
    End With
    
    wkb.Activate
    wkb.Sheets(1).Select
    ActiveWindow.WindowState = xlMaximized

ExitPoint:
    Set shtCategories = Nothing
    Set shtLambdas = Nothing
    Set loCategories = Nothing
    Set loLambdas = Nothing
    Set wkb = Nothing
    Set sht = Nothing
    StandardExit
    
End Sub



Sub ExportLambdaFunctionsFromActiveWorkbookToXml()

    Dim sXmlFileExportPath As String
    Dim sHumanReadableInventoryFilePath As String
    Dim wkb As Workbook
    Dim sExportPath As String
    Dim loLambdas As ListObject

    StandardEntry
    Set wkb = ActiveWorkbook

    If Not WorkbookIsValidForLambdaXmlExport(wkb) Then
        Exit Sub
    End If
    
    sExportPath = wkb.Path & Application.PathSeparator & "PowerFunctionExports"
    If Not FolderExists(sExportPath) Then CreateFolder (sExportPath)
    sXmlFileExportPath = sExportPath & Application.PathSeparator & "LambdaFunctions.xml"
    wkb.XmlMaps(gcsLambdaXmlMapName).Export URL:=sXmlFileExportPath, Overwrite:=True

    Set loLambdas = wkb.Sheets("Lambdas").ListObjects("tbl_Lambdas")
    sHumanReadableInventoryFilePath = sExportPath & Application.PathSeparator & "LambdaFunctions.txt"
    WriteHumanReadableLambdaInventory loLambdas, sHumanReadableInventoryFilePath

    MsgBox ("Functions exported")

    StandardExit

End Sub



Sub AddLambdaGitRepoToActiveWorkbook()

    Dim sRepoUrl As String
    Dim GitRepoStorage

    
    If ActiveWorkbook.Name = ThisWorkbook.Name Then Exit Sub
    
    sRepoUrl = InputBox("Enter Repo URL")
    If sRepoUrl = "" Then Exit Sub

    Set GitRepoStorage = AssignGitRepoStorage
    
    If RepoAlreadyExistsInStorage(sRepoUrl, GitRepoStorage) Then
        MsgBox ("This repo URL has previously been captured.  Current action ignored.")
    Else
        AddRepoToStorage sRepoUrl, GitRepoStorage
        MsgBox ("Repo successfully added")
    End If
    
    Set GitRepoStorage = Nothing

End Sub



Sub RefreshAvailableLambdaFormulas()

    Dim wkb As Workbook
    Dim LambdaStorage
    Dim LambdaFormulas As Dictionary
    Dim GitRepoStorage
    Dim sRepoList() As String
    Dim LambdaXmlMap As XmlMap

    StandardEntry
    Set wkb = ActiveWorkbook
    If wkb.Name = ThisWorkbook.Name Then Exit Sub


    If Not GitRepoStorageExists Then
        MsgBox ("There are no formula repos recorded in the active workbook")
    Else
        Set LambdaXmlMap = CreateLambdaXmlMap(wkb)
        Set GitRepoStorage = AssignGitRepoStorage
        Set LambdaStorage = AssignLambdaStorage
        ReadRepoList sRepoList, GitRepoStorage
        ImportDataIntoLambdaStorage sRepoList, LambdaStorage, LambdaXmlMap
        wkb.XmlMaps(gcsLambdaXmlMapName).Delete
        DeleteExistingLambdaFormulas wkb
        ReadLambdaFormulaDetails LambdaStorage, LambdaFormulas
        CreateLambdaFormulas wkb, LambdaFormulas
    End If

    Set GitRepoStorage = Nothing
    Set LambdaStorage = Nothing
    StandardExit

End Sub



Sub ShowLambdaUserForm()

    Dim LambdaStorage
    Dim uf As uf_LambdaFunctionWizard
    Dim i As Integer
    Dim LambdaNames


    StandardEntry
    Set LambdaStorage = AssignLambdaStorage
    Set uf = New uf_LambdaFunctionWizard
    Set uf.LambdaStorage = LambdaStorage

    uf.RefreshUserFormPropertiesFromStorage


    uf.Show
    Unload uf
    Set uf = Nothing
    Set LambdaStorage = Nothing
    StandardExit

End Sub




