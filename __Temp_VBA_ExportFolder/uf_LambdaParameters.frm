VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_LambdaParameters 
   Caption         =   "Lambda Parameters"
   ClientHeight    =   7220
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7780
   OleObjectBlob   =   "uf_LambdaParameters.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_LambdaParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "SpreadsheetBI"




Option Explicit

Private Type TypeUserFromLambdaParameters
    Params As Dictionary
    UserCancelled As Boolean
    DictionaryOfParamControls As Dictionary
End Type
Private this As TypeUserFromLambdaParameters


Public Property Get UserSelectedCancel() As Boolean
    UserSelectedCancel = this.UserCancelled
End Property


Public Property Let LambdaName(ByVal Name As String)
    Me.Caption = Name
End Property


Public Property Set ParameterDescriptions(ByVal Params As Dictionary)

    Dim key As Variant
    Dim ParamLabel As Control
    Dim ParamBox As TextBoxRangeSelector
    Dim i As Integer
    Const ParamLabelWidth As Integer = 125
    Const ParamTextBoxWidth As Integer = 125
    Const SpaceFromSides As Integer = 10
    Const SpaceBetweenParamTops As Integer = 30
    Const ParamHeight As Integer = 18
    
    Set this.Params = Params
    Set this.DictionaryOfParamControls = New Dictionary
    
    For i = 0 To Params.Count - 1
        
        Set ParamLabel = Me.FrameParameters.Controls.Add( _
            bstrProgID:="Forms.Label.1", _
            Name:="lbl_" & Params.Keys(i), _
            Visible:=True)
        ParamLabel.Width = ParamLabelWidth
        ParamLabel.Left = SpaceFromSides
        ParamLabel.Top = SpaceFromSides + (i * SpaceBetweenParamTops)
        ParamLabel.Height = ParamHeight
        ParamLabel.Caption = Params.Items(i)
        If ParamIsOptional(Params.Keys(i)) Then ParamLabel.Caption = _
            ParamLabel.Caption & " (optional)"
              
        Set ParamBox = New TextBoxRangeSelector
        ParamBox.Add Me.FrameParameters, "param_" & Params.Keys(i)
        ParamBox.Width = ParamTextBoxWidth
        ParamBox.Height = ParamHeight
        ParamBox.Left = (Me.FrameParameters.Width - ParamBox.Width - SpaceFromSides * 3)
        ParamBox.Top = SpaceFromSides + (i * SpaceBetweenParamTops)
        this.DictionaryOfParamControls.Add Item:=ParamBox, _
            key:=ParamBox.Name
                
    Next i
    
    If (Params.Count * SpaceBetweenParamTops) < Me.FrameParameters.Height Then
            Me.FrameParameters.ScrollBars = fmScrollBarsNone
    Else
        Me.FrameParameters.ScrollBars = fmScrollBarsVertical
        Me.FrameParameters.ScrollHeight = SpaceFromSides + (i * SpaceBetweenParamTops) + _
            (SpaceBetweenParamTops * 2)
    End If

End Property

Public Property Get OrderedParameterValues() As Variant

    Dim key As Variant
    Dim ReturnArray As Variant
    Dim i As Integer
    Dim ParamValue As String
    
    ReDim ReturnArray(0 To this.Params.Count - 1)
    
    For i = 0 To this.Params.Count - 1
        ParamValue = Me.FrameParameters.Controls("param_" & this.Params.Keys(i)).Value
        ReturnArray(i) = ParamValue
    Next i
        
    OrderedParameterValues = ReturnArray
    
End Property


Private Sub cbCancel_Click()
    this.UserCancelled = True
    Me.Hide
End Sub

Private Sub cbOk_Click()
    If ParametersAllPopulated Then
        Me.Hide
    Else
        MsgBox ("Please enter values for all non-optional parameters")
    End If
End Sub

Private Function ParametersAllPopulated() As Boolean

    Dim key As Variant
    Dim AllPopulated As Boolean
    
    AllPopulated = True
    For Each key In this.DictionaryOfParamControls.Keys
        AllPopulated = AllPopulated And _
            (this.DictionaryOfParamControls(key).Value <> "" Or ParamIsOptional(key))
    Next key
    ParametersAllPopulated = AllPopulated

End Function

Private Function ParamIsOptional(ByVal ParamName As String) As Boolean

    ParamIsOptional = Left(ParamName, 1) = "[" And Right(ParamName, 1) = "]"

End Function


