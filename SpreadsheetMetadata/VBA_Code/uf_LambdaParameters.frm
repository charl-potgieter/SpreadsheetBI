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
Option Explicit

Private Type TypeUserFromLambdaParameters
    Params As Dictionary
    UserCancelled As Boolean
End Type
Private this As TypeUserFromLambdaParameters


Public Property Get UserSelectedCancel() As Boolean
    UserSelectedCancel = this.UserCancelled
End Property


Public Property Let LambdaName(ByVal Name As String)
    Me.Caption = Name
End Property


Public Property Set ParameterDescriptions(ByVal Params As Dictionary)

    Dim Key As Variant
    Dim ParamLabel As Control
    Dim ParamRefBox As Control
    Dim i As Integer
    Const ParamLabelWidth As Integer = 125
    Const ParamRefBoxWidth As Integer = 125
    Const SpaceFromSides As Integer = 10
    Const SpaceBetweenParamTops As Integer = 30
    Const ParamHeight As Integer = 18
    
    Set this.Params = Params
    
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
              
        Set ParamRefBox = Me.FrameParameters.Controls.Add( _
            bstrProgID:="RefEdit.Ctrl", _
            Name:="re_" & Params.Keys(i), _
            Visible:=True)
        ParamRefBox.Width = ParamRefBoxWidth
        ParamRefBox.Height = ParamHeight
        ParamRefBox.Left = (Me.FrameParameters.Width - ParamRefBox.Width - SpaceFromSides * 3)
        ParamRefBox.Top = SpaceFromSides + (i * SpaceBetweenParamTops)
                
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

    Dim Key As Variant
    Dim ReturnArray As Variant
    Dim i As Integer
    
    ReDim ReturnArray(0 To this.Params.Count - 1)
    
    For i = 0 To this.Params.Count - 1
        ReturnArray(i) = Me.FrameParameters.Controls("re_" & this.Params.Keys(i)).Value
    Next i
        
    OrderedParameterValues = ReturnArray
    
End Property


Private Sub cbCancel_Click()
    this.UserCancelled = True
    Me.Hide
End Sub

Private Sub cbOK_Click()
    If ParametersAllPopulated Then
        Me.Hide
    Else
        MsgBox ("Please enter values for all non-optional parameters")
    End If
End Sub

Private Function ParametersAllPopulated() As Boolean

    Dim Key As Variant
    Dim AllPopulated As Boolean
    
    AllPopulated = True
    For Each Key In this.Params
        AllPopulated = AllPopulated And _
            (Me.FrameParameters.Controls("re_" & Key).Value <> "" Or ParamIsOptional(Key))
    Next Key
    ParametersAllPopulated = AllPopulated

End Function

Private Function ParamIsOptional(ByVal ParamName As String) As Boolean

    ParamIsOptional = Left(ParamName, 1) = "[" And Right(ParamName, 1) = "]"

End Function


Private Sub FrameParameters_Click()

End Sub
