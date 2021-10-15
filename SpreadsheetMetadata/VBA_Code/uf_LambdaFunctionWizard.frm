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


Private Sub cbGotoGitHub_Click()

    Dim LambdaNameSelected As String
    
    LambdaNameSelected = FirstListBoxSelection(Me.lbFunctions)
    Set this.LambdaDetailsSelected = this.LambdaFormulaDetails(LambdaNameSelected)

    ActiveWorkbook.FollowHyperlink this.LambdaDetailsSelected.GitUrl

End Sub

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


Private Sub lbFunctions_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
'Same behaviour as clicking ok
    
    Dim LambdaNameSelected
    LambdaNameSelected = FirstListBoxSelection(Me.lbFunctions)
    Set this.LambdaDetailsSelected = this.LambdaFormulaDetails(LambdaNameSelected)
    Me.Hide
    
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


Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
        this.UserCancelled = True
    End If
    
End Sub

