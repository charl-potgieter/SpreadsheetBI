Attribute VB_Name = "ZZZ_Testing"
Option Explicit


Sub TestResize()

    Dim rng As Range
    
    Set rng = ActiveSheet.Range(Columns(4), Columns(7))
    rng.Select

    Set rng = rng.Resize(rng.rows.Count - 3, rng.Columns.Count).Offset(3, 0)

    rng.Select

End Sub
