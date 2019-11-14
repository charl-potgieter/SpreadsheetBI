Attribute VB_Name = "m000_EntryPoints"
Option Explicit

Sub InsertFormattedSheetIntoActiveWorkbook()
    
    Dim sht As Worksheet
    
    Set sht = ActiveWorkbook.Sheets.Add(After:=ActiveSheet)
    FormatSheet sht

End Sub


Sub FormatZeroDecimalNumberFormat()

    SetNumberFormat "#,##0_);(#,##0);-??"

End Sub



Sub FormatOneDecimalNumberFormat()

    SetNumberFormat "#,##0.0_);(#,##0.0);-??"

End Sub



Sub FormatTwoDecimalsNumberFormat()

    
    SetNumberFormat "#,##0.00_);(#,##0.00);-??"
    
End Sub



Sub FormatTwoDigitPercentge()



End Sub
