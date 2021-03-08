Attribute VB_Name = "m005_MenuGenerator"
Option Explicit
Option Private Module


Sub DeletePopUpMenu()
'Delete PopUp menu if it exists
    
    On Error Resume Next
    Application.CommandBars(gcsMenuName).Delete
    On Error GoTo 0
    
End Sub



Sub CreatePopUpMenu()

    Dim cb As CommandBar
    Dim MenuCategory As CommandBarPopup
    Dim MenuSubcategory As CommandBarPopup
    Dim MenuItem As CommandBarControl
    
    Set cb = Application.CommandBars.Add(Name:=gcsMenuName, Position:=msoBarPopup, _
                                     MenuBar:=False, Temporary:=True)
    
    'Format main Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Format"
    
    
    'Format number submenu
    Set MenuSubcategory = MenuCategory.Controls.Add(Type:=msoControlPopup)
    MenuSubcategory.Caption = "Number"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Zero decimals"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatZeroDecimalNumberFormat"

    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "One decimal"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatOneDecimalNumberFormat"

    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Two decimals"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatTwoDecimalsNumberFormat"
                                     
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Zero decimals with arrows"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatZeroDecimalAndArrows"
                                     
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "One decimal with arrows"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatOneDecimalAndArrow"
                                     
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Two decimals with arrows"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatTwoDecimalsAndArrow"
                                     
                                     
    'Format percentage submenu
    Set MenuSubcategory = MenuCategory.Controls.Add(Type:=msoControlPopup)
    MenuSubcategory.Caption = "Percentage"
                                     
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Two digit percentage"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatTwoDigitPercentge"
                 
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Four digit percentage"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatFourDigitPercentge"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Zero digit percentage with arrows"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatZeroDigitPercentageAndArrow"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Two digit percentage with arrows"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatTwoDigitPercentageAndArrow"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Four digit percentage with arrows"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatFourDigitPercentageAndArrow"
    
    
    
    'Format dashboard submenu
    Set MenuSubcategory = MenuCategory.Controls.Add(Type:=msoControlPopup)
    MenuSubcategory.Caption = "Dashboard"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Dashboard icon style"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatDashboardIconStyle"
    
    
    
    'Format other submenu
    Set MenuSubcategory = MenuCategory.Controls.Add(Type:=msoControlPopup)
    MenuSubcategory.Caption = "Other"
       
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Date"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatDate"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "OK Error"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatOkError"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Active sheet"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatActiveSheet"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Active table"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatActiveTable"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Headings"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatHeadings"
    
    
    'Insert Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Insert"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Formatted sheet"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "InsertFormattedSheetIntoActiveWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Index page"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "InsertIndexPageActiveWorkbook"

    
    
    'Pivot reports Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Pivot reports"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Save report metadata"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "SavePivotReportMetadataInActiveWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Create report from metadata"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CreatePivotReportFromMetaData"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Create value copy"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CreateValueReportWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Create storage for supporting DAX queries"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "SetupPivotReportBackingStorage"
    
        
    'Pivot Table Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Pivot table"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Flatten"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatPivotTableFlatten"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Read pivot tables metadata"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ReadPivotTablesMetadata"
    
    
    'Data Model Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Data model"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Write data model info sheets"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "WriteModelInfoToSheets"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Copy power queries from another workbook"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CopyPowerQueriesFromWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Export power queries"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ExportPowerQueriesInActiveWorkbookToFiles"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Export non-standard power queries"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ExportNonStandardPowerQueriesInActiveWorkbookToFiles"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Import power queries in folder non-recursive"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ImportPowerQueriesFromSelectedFolderNonRecursive"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Import power queries in folder recursive"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ImportPowerQueriesFromSelectedFolderRecursive"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Import selected power queries"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ImportSelectedPowerQueries"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Create sheet for power query table generation"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CreatePowerQueryGeneratorSheet"
        
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Generate hard coded power query table"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "GeneratePowerQueryTable"
    
    
    'Other Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Other"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Run looper"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "TableLooper"
    
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Export all tables in workbook as pipe delimited text"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ExportTablesInActiveWorkbookToPipeDelimtedText"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Generate spreadsheet metadata"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "GenerateSpreadsheetMetaData"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Export VBA Code excluding module name"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ExportVBAcodeExModuleName"
    
    

    
    
End Sub

