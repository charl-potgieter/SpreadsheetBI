VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextBoxRangeSelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'---------------------------------------------------------------------------------
' Implements a textbox / inputbox combo to avoid issues with refedit
' Inspired by https://peltiertech.com/refedit-control-alternative/
'---------------------------------------------------------------------------------


Public WithEvents EventTextBox As MSForms.TextBox
Attribute EventTextBox.VB_VarHelpID = -1

Private Type TypeTextBoxRangeSelector
    TextBox As Control
    TextBoxResidesOn As Object 'represents Userform / Frame / Multipage etc where the textbox will be placed
End Type

Private this As TypeTextBoxRangeSelector



Public Sub Add(ByVal TextBoxResidesOn, ByVal Name As String)

'TextBoxResidesOn represents Userform / Frame / Multipage etc where the textbox will be placed
        
    Set this.TextBox = TextBoxResidesOn.Controls.Add( _
        bstrProgID:="Forms.TextBox.1", _
        Name:=Name, _
        Visible:=True)

    this.TextBox.DropButtonStyle = fmDropButtonStyleReduce
    this.TextBox.ShowDropButtonWhen = fmShowDropButtonWhenAlways
    Set this.TextBoxResidesOn = TextBoxResidesOn
    
    Set Me.EventTextBox = this.TextBox

End Sub
    

Public Property Get Name() As String
    Name = this.TextBox.Name
End Property


Public Property Get Value() As Variant
    Value = this.TextBox.Text
End Property

Public Property Let Width(tbWidth As Double)
    this.TextBox.Width = tbWidth
End Property


Public Property Get Width() As Double
    Width = this.TextBox.Width
End Property


Public Property Let Height(tbHeight As Double)
    this.TextBox.Height = tbHeight
End Property


Public Property Let Left(tbLeft As Double)
    this.TextBox.Left = tbLeft
End Property


Public Property Let Top(tbTop As Double)
    this.TextBox.Top = tbTop
End Property


Private Sub EventTextBox_DropButtonClick()
    
    Dim UserInput As Variant
    Dim ScreenUpdatingIsDisabledAtEntry  As Boolean
    Dim UnstructuredRange As String
  
    'Enable screen updating to see the "marching ants" around range selected
    If Application.ScreenUpdating Then
        ScreenUpdatingIsDisabledAtEntry = False
    Else
        ScreenUpdatingIsDisabledAtEntry = True
        Application.ScreenUpdating = True
    End If
    
    UserInput = Application.InputBox(Prompt:="Select the range", _
        Title:="Select range or enter value", _
        Left:=this.TextBoxResidesOn.Left, Top:=this.TextBoxResidesOn.Top, Type:=0)
    
    'Remove the equals sign, change to relative reference and convert from R1C1 to A1-style
    'referencing when setting value in textbox
    UnstructuredRange = Replace( _
        Application.ConvertFormula( _
            Formula:=Right(UserInput, Len(UserInput) - 1), _
            FromReferenceStyle:=xlR1C1, ToReferenceStyle:=xlA1), _
        "$", "")
    
    this.TextBox.Value = ConvertReferenceToStructuredReference(ActiveCell, UnstructuredRange)
    
    
    If ScreenUpdatingIsDisabledAtEntry Then
        Application.ScreenUpdating = False
    End If


End Sub


Private Sub EventTextBox_Enter()
    MsgBox ("Entered")
End Sub
