VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LambdaFormulaDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "SpreadsheetBI"
Option Explicit

'Cannot add a type to a dictionary hence use a class
' https://stackoverflow.com/questions/9178177/how-to-put-user-defined-datatype-into-a-dictionary

Public Name As String
Public RefersTo As String
Public Category As String
Public Author As String
Public Description As String
Public ParameterDescriptions As Dictionary
Public GitUrl As String
