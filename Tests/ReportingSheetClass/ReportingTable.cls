VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReportingTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'@Folder "Reporting"
Option Explicit

Private Type TReportingTable
    tblReport As ListObject
End Type

Private this As TReportingTable

Public Property Get Name() As String
   Name = this.tblReport.Name
End Property

Public Property Let Name(sName As String)
    this.tblReport.Name = sName
End Property


Public Sub Create(ByVal rng As Range)
    Dim sht As Worksheet
    Set sht = rng.Parent
    sht.ListObjects.Add xlSrcRange, rng.Cells(1)
End Sub


Public Sub SetDaxQuerySource(ByVal sQuery As String)

End Sub
