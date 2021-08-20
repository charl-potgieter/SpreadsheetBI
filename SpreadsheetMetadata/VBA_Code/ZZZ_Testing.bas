Attribute VB_Name = "ZZZ_Testing"
Option Explicit


Sub TestResize()

    Dim rng As Range
    
    Set rng = ActiveSheet.Range(Columns(4), Columns(7))
    rng.Select

    Set rng = rng.Resize(rng.Rows.Count - 3, rng.Columns.Count).Offset(3, 0)

    rng.Select

End Sub




Sub LoopThroughSelectedSheets()
'To loop through all selected sheets in
'the active workbook
   
   Dim sh As Object
   Dim sht As Worksheet
   
   For Each sh In ActiveWindow.SelectedSheets
    Set sht = sh
   
     Debug.Print sht.Name
   Next sh
End Sub



