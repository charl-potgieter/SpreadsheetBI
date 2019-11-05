Attribute VB_Name = "m000_EntryPoints"
Option Explicit

Sub InsertFormattedSheetIntoActiveWorkbook()
    
    Dim sht As Worksheet
    
    Set sht = ActiveWorkbook.Sheets.Add(After:=ActiveSheet)
    FormatSheet sht

End Sub



