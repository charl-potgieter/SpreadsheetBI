VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_LambdaFunctionWizard 
   Caption         =   "Insert Power Function"
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
    LambdaStorage As ListStorage
    LambdaFormulaDetails As Dictionary
    EventsAreEnabled As Boolean
    UserCancelled As Boolean
End Type
Private this As TypePowerFunctionWizard


Public Property Set LambdaStorage(ByRef Storage)
'cannot pass variables to  userform event so store as a class property (userforms are classes)
    Set this.LambdaStorage = Storage
End Property


Public Property Get UserSelectedCancel() As Boolean
    UserSelectedCancel = this.UserCancelled
End Property


Property Let EnableEvents(Enable As Boolean)

End Property


Sub RefreshUserFormPropertiesFromStorage()

    Dim i As Integer
    Dim Categories
    
    ReadUniqueLambdaCategories this.LambdaStorage, Categories
    Me.comboCategories.AddItem "All"
    For i = LBound(Categories) To UBound(Categories)
        Me.comboCategories.AddItem Categories(i)
    Next i
    Me.comboCategories.Value = "All"
    
    ReadLambdaFormulaDetails this.LambdaStorage, this.LambdaFormulaDetails

End Sub



Private Sub comboCategories_Change()

    Dim LambdaNamesPerCategorySelection
    Dim i As Integer
    
    ReadLambdaNamesPerCategory this.LambdaStorage, LambdaNamesPerCategorySelection, Me.comboCategories.Value
    
    Me.lbFunctions.Clear
    For i = LBound(LambdaNamesPerCategorySelection) To UBound(LambdaNamesPerCategorySelection)
        Me.lbFunctions.AddItem LambdaNamesPerCategorySelection(i)
    Next i

End Sub




Private Sub lbFunctions_Click()
    
    Dim DisplayString As String
    Dim FormulaNameSelected As String
    Dim FormulaDetail As LambdaFormulaDetails
    Dim ParameterDescriptions As Dictionary
    Dim Parameter
    Dim IsFirstParamter As Boolean
    
    FormulaNameSelected = Me.lbFunctions.Value
    Set FormulaDetail = this.LambdaFormulaDetails(FormulaNameSelected)
    Set ParameterDescriptions = FormulaDetail.ParameterDescriptions
        
    DisplayString = FormulaNameSelected & "("
    
    IsFirstParamter = True
    For Each Parameter In ParameterDescriptions.Keys
        If IsFirstParamter Then
            DisplayString = DisplayString & Parameter
        Else
            DisplayString = DisplayString & ", " & Parameter
        End If
    Next Parameter
    DisplayString = DisplayString & ")"
    
    
    If Len(DisplayString) > 100 Then
        DisplayString = Left(DisplayString, 98) & "..."
    End If

    'me.tbFunction.Text.Font =
    Me.tbFunction.Value = DisplayString
    
    Me.tbDescription = FormulaDetail.Description
    
    
End Sub

Private Sub UserForm_Terminate()
   Set this.LambdaStorage = Nothing
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
        this.UserCancelled = True
    End If
    
End Sub
