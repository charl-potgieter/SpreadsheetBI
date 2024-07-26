Attribute VB_Name = "ZZZ_Testing"
Option Explicit



Sub GroupColumn()

    ActiveSheet.Range("A:B").Columns.Group


End Sub




Sub Showdetails()


    If ActiveSheet.Columns("A").Hidden Then
        ActiveSheet.Range("A:B").Resize(, 1).EntireColumn.ShowDetail = True
    Else
        ActiveSheet.Range("A:B").Resize(, 1).EntireColumn.ShowDetail = False
    End If


End Sub


