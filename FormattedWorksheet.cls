VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FormattedWorksheet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'@Folder "Formatting"
'@PredeclaredId
Option Explicit

Private Type TypeFormattedWorksheet
    Heading As String
    Category As String
    FreezePaneCol As Long
    FreezePaneRow As Long
End Type


Private This As TypeFormattedWorksheet

Private Sub Class_Initialize()
    MsgBox ("Hello")
End Sub


Property Let SheetHeading(ByVal Heading As String)
    This.Heading = Heading
End Property
