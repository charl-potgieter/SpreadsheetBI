Attribute VB_Name = "m999_TestDmvQuery"
Option Explicit
Option Private Module


Public Sub CreatePowerPivotDmvInventory()
'Requires reference to Microsoft ActiveX Data Objects

    Dim conn As ADODB.Connection
    Dim sht As Excel.Worksheet
    Dim iRowNum As Integer
    Dim i As Integer
    Dim lo As ListObject

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False


    ' Open connection to PowerPivot engine
    Set conn = ActiveWorkbook.Model.DataModelConnection.ModelConnection.ADOConnection
    
    'Create output sheet
    If SheetExists(ActiveWorkbook, "DMV") Then
        ActiveWorkbook.Sheets("DMV").Delete
    End If
    Set sht = ActiveWorkbook.Worksheets.Add
    sht.Name = "DMV"
    
    Set lo = ThisWorkbook.Sheets("DMV_Names").ListObjects("tbl_DMV_names")
    
    ' Call function by passing the DMV name
    iRowNum = 1
    With lo
        For i = 1 To .DataBodyRange.rows.Count
            WriteDmvContent .ListColumns("DMV Name").DataBodyRange.Cells(i), conn, sht, iRowNum
            'Application.Wait Now + #12:00:01 AM#
       Next i
    End With

    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True


End Sub


Private Sub WriteDmvContent(ByVal dmvName As String, ByRef conn As ADODB.Connection, ByRef Sheet As Worksheet, ByRef iRowNum)

    Dim rs As ADODB.Recordset
    Dim cell As Excel.Range
    Dim sSQL As String
    Dim i As Integer
    
    ' SQL like query to get result of DMV from schema $SYSTEM
    sSQL = "select * from $SYSTEM." & dmvName
    
    Set rs = New ADODB.Recordset
    rs.ActiveConnection = conn
    
    On Error GoTo ExitPoint
    rs.Open sSQL, conn, adOpenForwardOnly, adLockOptimistic
    On Error GoTo 0
    
    ' Output of the DMV name
    Sheet.Cells(iRowNum, 1) = dmvName
    FormatDmvName Sheet.Cells(iRowNum, 1)
    
    iRowNum = iRowNum + 1
    ' Output of the column names
    For i = 0 To rs.Fields.Count - 1
        Sheet.Cells(iRowNum, i + 1) = rs.Fields(i).Name
        FormatColumnHeader Sheet.Cells(iRowNum, i + 1)
    Next i
    
    iRowNum = iRowNum + 1
    ' Output of the query results
    Do Until rs.EOF
        For i = 0 To rs.Fields.Count - 1
            WriteFormatedCellValue Sheet.Cells(iRowNum, i + 1), rs.Fields(i)
        Next i
    
        iRowNum = iRowNum + 1
        rs.MoveNext
    Loop
    
    rs.Close
    
    iRowNum = iRowNum + 2

ExitPoint:
    Set rs = Nothing

End Sub


Private Sub WriteFormatedCellValue(ByRef cell As Excel.Range, ByRef field As ADODB.field)
' Write and Format the value depending on the data type

    On Error GoTo ExitPoint
    cell = field.Value
    
    Select Case field.Type
        
        Case ADODB.DataTypeEnum.adDate, DataTypeEnum.adDBDate
            cell.NumberFormat = "m/d/yyyy"
        
        Case DataTypeEnum.adBigInt, DataTypeEnum.adInteger, DataTypeEnum.adSmallInt, DataTypeEnum.adTinyInt
            cell.NumberFormat = "###0"
        
        Case DataTypeEnum.adCurrency, DataTypeEnum.adDecimal, DataTypeEnum.adDouble, DataTypeEnum.adNumeric, DataTypeEnum.adSingle
            cell.NumberFormat = "#,##0.00"
    End Select
    
ExitPoint:

End Sub


Private Sub FormatDmvName(ByRef cell As Excel.Range)
' Write the formatted DMV name to the sheet.

    cell.Font.Bold = True
    With cell.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .ThemeColor = xlThemeColorDark1
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

End Sub


Private Sub FormatColumnHeader(ByRef cell As Excel.Range)
' Write the formatted DMV column name to the sheet.

    cell.Font.Bold = True
    With cell.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 14408667
        .TintAndShade = -0.249977111117893
        .PatternTintAndShade = 0
    End With

End Sub

