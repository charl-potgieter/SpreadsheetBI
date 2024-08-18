Attribute VB_Name = "ZZZ_Testing"
Option Explicit


Function SheetErrorStatusFormula2() As String



    SheetErrorStatusFormula2 = _
        "=LET(" & vbLf & _
        "    ErrorChecksOK, AND(TRUE, ErrorCheckRows, ErrorCheckColumns)," & vbLf & _
        "    ErrorOnErrorCheck, ISERROR(ErrorChecksOK)," & vbLf & _
        "    SWITCH(" & vbLf & _
        "        TRUE," & vbLf & _
        "        ErrorOnErrorCheck, ""Sheet error - see ranges ErrrorCheckColumns and ErrorCheckRows""," & vbLf & _
        "        NOT(ErrorChecksOK), ""Sheet error - see ranges ErrrorCheckColumns and ErrorCheckRows""," & vbLf & _
        "        COUNTIFS(Index!HiddenCategoriesCol, Category, Index!ReportNamesCol, Heading) = 0, ""This sheet heading / category combination does not appear on index tab""," & vbLf & _
        "        COUNTIFS(Index!HiddenCategoriesCol, Category, Index!ReportNamesCol, Heading) > 1, ""This sheet heading / category combination appears multiple times on index tab""," & vbLf & _
        "        ""OK""" & vbLf & _
        "    )" & vbLf & _
        ")"


End Function
