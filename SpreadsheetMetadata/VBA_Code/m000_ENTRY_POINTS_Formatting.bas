Attribute VB_Name = "m000_ENTRY_POINTS_Formatting"
Option Explicit


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
    SetNumberFormat "0.00%"
End Sub


Sub FormatFourDigitPercentge()
    SetNumberFormat "0.0000%"
End Sub


Sub FormatDate()
    SetNumberFormat "dd-mmm-yy"
End Sub


Sub FormatDashboardIconStyle()
'Creates custom formatting to displat different dashboard style icons, for positive, negative, zero and text values
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color 10] " & ChrW(&H25B2) & "_);" & _
        "[Red] " & ChrW(&H25BC) & "_);" & _
        "[Color 46] " & ChrW(&H2666) & " ;" & _
        "[Blue] * " & ChrW(&H25BA) & "_ "

End Sub


Sub FormatZeroDecimalAndArrows()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10]#,##0_) " & ChrW(&H25B2) & "_);" & _
        "[Red] (#,##0) " & ChrW(&H25BC) & "_);" & _
        "-????;" & _
        "General"

End Sub

Sub FormatOneDecimalAndArrow()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10]#,##0.0_) " & ChrW(&H25B2) & "_);" & _
        "[Red] (#,##0.0) " & ChrW(&H25BC) & "_);" & _
        "-????;" & _
        "General"

End Sub

Sub FormatTwoDecimalsAndArrow()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10]#,##0.00_) " & ChrW(&H25B2) & "_);" & _
        "[Red] (#,##0.00) " & ChrW(&H25BC) & "_);" & _
        "-????;" & _
        "General"

End Sub

Sub FormatZeroDigitPercentageAndArrow()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10] 0% " & ChrW(&H25B2) & "_);" & _
        "[Red] -0% " & ChrW(&H25BC) & "_);" & _
        "0%??;" & _
        "General"

End Sub

Sub FormatTwoDigitPercentageAndArrow()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10] 0.00% " & ChrW(&H25B2) & "_);" & _
        "[Red] -0.00% " & ChrW(&H25BC) & "_);" & _
        "0.00%??;" & _
        "General"

End Sub

Sub FormatFourDigitPercentageAndArrow()
'Custom formatting displays numbers and up and down arrows as appropriate
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Note that Hex character codes are obtained by using excel menu, insert -> symbol
'(select font as arial top right, subset as geometric shape, select hex code bottom left.
'Note also use of _) and * below to get good alignment of symbols (the bracket seems to give enough space to align triangles with diamond)
'Sheet Font should be set to Calibri to ensure best display (which is as per the m010_General.FormatSheet sub in this workbook)
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10] 0.0000% " & ChrW(&H25B2) & "_);" & _
        "[Red] -0.0000% " & ChrW(&H25BC) & "_);" & _
        "0.0000%??;" & _
        "General"

End Sub

Sub FormatOkError()
'1 Displays OK in green, zero ERROR in red.  Negatives adn text are hidden
'Can be used in conjuction with Power Pivot measures  designed to generate numbers to achieve desired icon style
'Useful links and inspiration
'   https://www.youtube.com/watch?v=tGY70sdpaLc&t=14s
'   https://www.xelplus.com/smart-uses-of-custom-formatting/
'   https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2007/cc296089(v=office.12)?redirectedfrom=MSDN

    SetNumberFormat "[Color10]OK ;;[Red]\E\R\RO\R;"

End Sub


Sub FormatActiveSheet()

    FormatSheet ActiveSheet

End Sub



Sub FormatHeadings()

    StandardEntry

    'Remove all current borders
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone

    'Set new borders
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With

    'Set header colour
    With Selection.Interior
        .Color = RGB(217, 225, 242)
        .Pattern = xlSolid
    End With

    Selection.Font.Bold = True

    'Set Text allignment
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlTop
        .WrapText = True
    End With

    StandardExit

End Sub






Sub FormatActiveTable()

    FormatTable ActiveCell.ListObject

End Sub

