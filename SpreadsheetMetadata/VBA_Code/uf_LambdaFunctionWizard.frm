VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_LambdaFunctionWizard 
   Caption         =   "Lambda function wizard"
   ClientHeight    =   6970
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8780.001
   OleObjectBlob   =   "uf_LambdaFunctionWizard.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_LambdaFunctionWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


'Use "This" declaration as an easy way to get intellisense to the classes private variables
'https://rubberduckvba.wordpress.com/2020/02/27/vba-classes-gateway-to-solid/
Private Type TypePowerFunctionWizard
    Categories As Variant
    LambdaFormulaDetails As Dictionary
    UserCancelled As Boolean
    LambdaDetailsSelected As LambdaFormulaDetails
End Type
Private this As TypePowerFunctionWizard


Private Sub cbOK_Click()

    Dim LambdaNameSelected
    
    LambdaNameSelected = FirstListBoxSelection(Me.lbFunctions)
    Set this.LambdaDetailsSelected = this.LambdaFormulaDetails(LambdaNameSelected)
    Me.Hide

End Sub


Private Sub cbCancel_Click()
    Me.Hide
    this.UserCancelled = True
End Sub


Private Sub cbGetFormulaText_Click()
    
    Dim sht As Worksheet
    
    Set sht = ActiveWorkbook.Sheets.Add
    FormatSheet sht
    
    With sht.Range("B5")
       
    
    End With
      
'        Application.CutCopyMode = False
'    Selection.NumberFormat = "@"
'    With Selection
'        .HorizontalAlignment = xlGeneral
'        .VerticalAlignment = xlBottom
'        .WrapText = False
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
'    With Selection
'        .HorizontalAlignment = xlLeft
'        .VerticalAlignment = xlTop
'        .WrapText = True
'        .Orientation = 0
'        .AddIndent = False
'        .IndentLevel = 0
'        .ShrinkToFit = False
'        .ReadingOrder = xlContext
'        .MergeCells = False
'    End With
'    Columns("B:B").EntireColumn.AutoFit
'    Rows("2:2").EntireRow.AutoFit
'    Columns("B:B").ColumnWidth = 96.91
'    Rows("2:2").EntireRow.AutoFit
    Columns("B:B").Select
    
    
End Sub



Private Sub comboCategories_Change()

    Dim Key As Variant
    
    Me.lbFunctions.Clear
    For Each Key In this.LambdaFormulaDetails.Keys
        If Me.comboCategories.Value = "All" Or _
            this.LambdaFormulaDetails(Key).Category = Me.comboCategories.Value Then
                Me.lbFunctions.AddItem Key
        End If
    Next Key
    Me.lbFunctions.ListIndex = 0
    
End Sub



Private Sub lbFunctions_Change()
    DisplayLambdaNameParametersAndDescription
End Sub


Private Sub UserForm_Initialize()
    this.UserCancelled = False
End Sub

Private Sub UserForm_Terminate()
   Set this.LambdaFormulaDetails = Nothing
   Set this.LambdaDetailsSelected = Nothing
End Sub


Public Property Get UserSelectedCancel() As Boolean
    UserSelectedCancel = this.UserCancelled
End Property


Public Property Let Categories(ByVal LambdaCategories As Variant)
    
    Dim i As Integer
    
    this.Categories = LambdaCategories
    Me.comboCategories.Clear
    Me.comboCategories.AddItem "All"
    For i = LBound(LambdaCategories) To UBound(LambdaCategories)
        Me.comboCategories.AddItem LambdaCategories(i)
    Next i
    Me.comboCategories.ListIndex = 0
    
End Property

Public Property Set LambdaDetails(ByVal LambdaDetailDict As Dictionary)
    
    Dim Key As Variant
    
    Set this.LambdaFormulaDetails = LambdaDetailDict
    For Each Key In this.LambdaFormulaDetails.Keys
        Me.lbFunctions.AddItem Key
    Next Key
    Me.lbFunctions.ListIndex = 0
    DisplayLambdaNameParametersAndDescription
    
End Property




Public Property Get SelectedLambdaDetails() As LambdaFormulaDetails
    Set SelectedLambdaDetails = this.LambdaDetailsSelected
End Property



Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
        this.UserCancelled = True
    End If
    
End Sub


Private Sub DisplayLambdaNameParametersAndDescription()
    
    Dim LambdaName As String
    Dim LambdaDetail As LambdaFormulaDetails
    Dim LambaDisplayString As String
    Dim DescriptionDisplayString As String
    Dim Key As Variant
    Dim IsFirstKeyInDictionary As Boolean

    If NumberOfSelectedItemsInListBox(Me.lbFunctions) = 1 Then
        LambdaName = Me.lbFunctions.Value
        Set LambdaDetail = this.LambdaFormulaDetails(LambdaName)
        LambaDisplayString = LambdaName & "("
        IsFirstKeyInDictionary = True
        For Each Key In LambdaDetail.ParameterDescriptions.Keys
            If Not IsFirstKeyInDictionary Then LambaDisplayString = LambaDisplayString & ", "
            IsFirstKeyInDictionary = False
            LambaDisplayString = LambaDisplayString & Key
        Next Key
        LambaDisplayString = LambaDisplayString & ")"
        DescriptionDisplayString = LambdaDetail.Description
    Else
        LambaDisplayString = ""
        DescriptionDisplayString = ""
    End If

    Me.tbFunction.Value = LambaDisplayString
    Me.tbDescription = DescriptionDisplayString

End Sub



'Public Property Set LambdaStorage(ByRef Storage)
''cannot pass variables to  userform event so store as a class property (userforms are classes)
'    Set this.LambdaStorage = Storage
'End Property






'
'Sub RefreshUserFormPropertiesFromStorage()
'
'    Dim i As Integer
'    Dim Categories
'
'    ReadUniqueLambdaCategories this.LambdaStorage, Categories
'    Me.comboCategories.AddItem "All"
'    For i = LBound(Categories) To UBound(Categories)
'        Me.comboCategories.AddItem Categories(i)
'    Next i
'    Me.comboCategories.Value = "All"
'
'    ReadLambdaFormulaDetails this.LambdaStorage, this.LambdaFormulaDetails
'
'End Sub
'
'
'
'Private Sub comboCategories_Change()
'
'    Dim LambdaNamesPerCategorySelection
'    Dim i As Integer
'
'    ReadLambdaNamesPerCategory this.LambdaStorage, LambdaNamesPerCategorySelection, Me.comboCategories.Value
'
'    Me.lbFunctions.Clear
'    For i = LBound(LambdaNamesPerCategorySelection) To UBound(LambdaNamesPerCategorySelection)
'        Me.lbFunctions.AddItem LambdaNamesPerCategorySelection(i)
'    Next i
'
'End Sub
'
'
'
'
'Private Sub lbFunctions_Click()
'
'    Dim LambaDisplayString As String
'    Dim FormulaNameSelected As String
'    Dim FormulaDetail As LambdaFormulaDetails
'    Dim ParameterDescriptions As Dictionary
'    Dim Parameter
'    Dim IsFirstParamter As Boolean
'
'    FormulaNameSelected = Me.lbFunctions.Value
'    Set FormulaDetail = this.LambdaFormulaDetails(FormulaNameSelected)
'    Set ParameterDescriptions = FormulaDetail.ParameterDescriptions
'
'    LambaDisplayString = FormulaNameSelected & "("
'
'    IsFirstParamter = True
'    For Each Parameter In ParameterDescriptions.Keys
'        If IsFirstParamter Then
'            LambaDisplayString = LambaDisplayString & Parameter
'        Else
'            LambaDisplayString = LambaDisplayString & ", " & Parameter
'        End If
'    Next Parameter
'    LambaDisplayString = LambaDisplayString & ")"
'
'
'    If Len(LambaDisplayString) > 100 Then
'        LambaDisplayString = Left(LambaDisplayString, 98) & "..."
'    End If
'
'    'me.tbFunction.Text.Font =
'    Me.tbFunction.Value = LambaDisplayString
'
'    Me.tbDescription = FormulaDetail.Description
'
'
'End Sub





