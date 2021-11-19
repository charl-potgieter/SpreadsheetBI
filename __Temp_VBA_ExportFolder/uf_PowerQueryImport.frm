VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} uf_PowerQueryImport 
   Caption         =   "Power query import"
   ClientHeight    =   7170
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   8060
   OleObjectBlob   =   "uf_PowerQueryImport.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "uf_PowerQueryImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit


Private Type TypeUserFormPowerQueryImport
    pqDetailsAll As Dictionary
    pqDetailsSelected As Dictionary
    UserCancelled As Boolean
End Type
Private this As TypeUserFormPowerQueryImport


Public Property Set PowerQueryDetails(pqDetailsAll As Dictionary)

    Dim UniqueCategories As Collection
    Dim QueryName As Variant
    Dim Category As Variant
    
    Set this.pqDetailsAll = pqDetailsAll
    
    'Get unique categories
    Set UniqueCategories = New Collection
    For Each QueryName In pqDetailsAll.Keys
        On Error Resume Next
        UniqueCategories.Add item:=pqDetailsAll(QueryName).Category, _
            key:=pqDetailsAll(QueryName).Category
        On Error GoTo 0
    Next QueryName
    
    BubbleSortSortCollection UniqueCategories
    
    Me.comboCategories.AddItem "All"
    For Each Category In UniqueCategories
        Me.comboCategories.AddItem Category
    Next Category

End Property



Public Property Get UserSelectedCancel() As Boolean
    UserSelectedCancel = this.UserCancelled
End Property


Public Property Get SelectedQueries() As Dictionary

    Set SelectedQueries = this.pqDetailsSelected

End Property


Private Sub comboCategories_Change()

    Dim key As Variant

    If Me.comboCategories.Value = "All" Then
        For Each key In this.pqDetailsAll.Keys
            Me.lbQueries.AddItem key
        Next key
    Else
        Me.lbQueries.Clear
        For Each key In this.pqDetailsAll.Keys
            If this.pqDetailsAll.item(key).Category = Me.comboCategories.Value Then
                Me.lbQueries.AddItem key
            End If
        Next key
    End If
    Me.lbQueries.Selected(0) = True

End Sub


Private Sub lbQueries_Change()

    Dim SingleSelectedListboxValue As String
    Dim QueryDescription As String
'
    If NumberOfSelectedItemsInListBox(lbQueries) = 1 Then
        SingleSelectedListboxValue = ArrayOfListBoxSelections(lbQueries)(0)
        Me.tblDescription = this.pqDetailsAll.item(SingleSelectedListboxValue).Description
    Else
        Me.tblDescription = ""
    End If

End Sub


Private Sub UserForm_Initialize()
    this.UserCancelled = False
    Set this.pqDetailsSelected = Nothing
End Sub


Private Sub cbOk_Click()

    Dim i As Long
    Set this.pqDetailsSelected = New Dictionary
    
    If UserFormListBoxHasSelectedItems(Me.lbQueries) Then
        Set this.pqDetailsSelected = New Dictionary
        For i = 0 To Me.lbQueries.ListCount - 1
           If Me.lbQueries.Selected(i) Then
                this.pqDetailsSelected.Add key:=Me.lbQueries.List(i), _
                    item:=this.pqDetailsAll(Me.lbQueries.List(i))
           End If
        Next i
    End If

    Me.Hide
End Sub


Private Sub cbCancel_Click()
    Me.Hide
    this.UserCancelled = True
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer _
                                       , CloseMode As Integer)
    
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Me.Hide
        this.UserCancelled = True
    End If
    
End Sub

Private Sub UserForm_Terminate()
    Set this.pqDetailsAll = Nothing
    Set this.pqDetailsSelected = Nothing
End Sub
