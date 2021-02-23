VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PowerReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'Use "This" declaration as an easy way to get intellisense to the classes private variables
'https://rubberduckvba.wordpress.com/2020/02/27/vba-classes-gateway-to-solid/
Private Type TypePowerReport
    SheetName As String
    SheetHeading As String
    SheetCategory As String
    PvtTable As PivotTable
End Type

Private This As TypePowerReport


Public Property Get SheetName() As String
    SheetName = This.SheetName
End Property


Public Property Get SheetHeading() As String
    SheetHeading = This.SheetHeading
End Property


Public Property Let SheetCategory(sSheetCategory As String)
    This.SheetCategory = sSheetCategory
End Property


Public Property Get SheetCategory() As String
    SheetCategory = This.SheetCategory
End Property

Public Property Set PvtTable(pvt As PivotTable)
    Set This.PvtTable = pvt
End Property


Public Property Get PvtTable() As PivotTable
    Set PvtTable = This.PvtTable
End Property


Public Sub GetSheetProperties(ByRef sht As Worksheet)
    This.SheetName = sht.Name
    This.SheetHeading = sht.Range("SheetHeading")
    This.SheetCategory = sht.Range("SheetCategory")
End Sub

