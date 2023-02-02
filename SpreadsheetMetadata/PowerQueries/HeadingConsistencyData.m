let
    ExcelData = fn_HeadingConsistency("D:\Onedrive\Documents_Charl\Computer_Technical\Programming_GitHub\SpreadsheetBI\Testing\Test_HeadingConsistency\Test_ExcelData","Excel Data"),
    ExcelTable = fn_HeadingConsistency("D:\Onedrive\Documents_Charl\Computer_Technical\Programming_GitHub\SpreadsheetBI\Testing\Test_HeadingConsistency\Test_ExcelTable","Excel Table"),
    CSV_Data= fn_HeadingConsistency("D:\Onedrive\Documents_Charl\Computer_Technical\Programming_GitHub\SpreadsheetBI\Testing\Test_HeadingConsistency\Test_CSV","CSV"),
    Custom = fn_HeadingConsistency("D:\Onedrive\Documents_Charl\Computer_Technical\Programming_GitHub\SpreadsheetBI\Testing\Test_HeadingConsistency\Test_ExcelCustom","other", HeadingConsisency_TestCustomFunction),

    CombinedTables = Table.Combine({ExcelData, ExcelTable,CSV_Data, Custom})
in
    CombinedTables