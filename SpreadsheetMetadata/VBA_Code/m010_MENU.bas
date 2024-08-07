Attribute VB_Name = "m010_MENU"
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
    MenuItem.Caption = "Active Sheet (simple format, not reporting sheet)"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatActiveSheet"
       
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Date"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatDate"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "OK Error"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatOkError"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Active table"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatActiveTable"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Headings"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatHeadings"
    
    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Update default report sheet format"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "SetReportSheetFormat"
    
'    Set MenuItem = MenuSubcategory.Controls.Add(Type:=msoControlButton)
'    MenuItem.Caption = "Save default report sheet format"
'    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "SaveReportSheetFormat"
    
    
    'Insert Menu
    
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Insert"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Index page"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "InsertIndexPageActiveWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Reporting sheet"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "InsertReportingSheetSheetIntoActiveWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "List storage sheet"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "InsertStorageSheet"
    
    
    'Reports Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Reports"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Save report metadata"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "SaveReportMetadataInActiveWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Create report from metadata"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CreateReportFromMetadata"
    
'    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
'    MenuItem.Caption = "Create storage for table reports"
'    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "AssignTableReportStorageInActiveWorkbook"
    
'    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
'    MenuItem.Caption = "Create storage for Excel table formulas"
'    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "AssignPivotReportFormulaStorageInActiveWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Create storage for Dax queries per report"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "AssignPivotReportQueriesPerReportActiveWorkbook"
     
     
    'Pivot Table Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Pivot table"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Flatten"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "PivotTableFlatten"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Read pivot tables metadata"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ReadPivotTablesMetadata"
    
    
    'Power Query Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Power Queries"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Import or refresh Power Query Library"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ImportPowerQueryLibraryIntoActiveWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Create consolidated power query"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CreateConsolidatedPowerQuery"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Copy power queries from another workbook"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CopyPowerQueriesFromWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Export power queries"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ExportPowerQueriesInActiveWorkbookToFiles"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Export consolidated power queries"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ExportPowerQueriesInActiveWorkbookToConsolidatedFile"
    
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
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Create power queries referenced to text files"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CreateTextReferencedPowerQueriesInActiveWorkbook"
    
        
    'Data Model Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Data model"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Write data model info sheets"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "WriteModelInfoToSheets"
    
    
    
    'Lambda menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Lambdas"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Insert lambda function"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "InsertLambda"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Refresh library from GitHub"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "RefreshLambdaLibraries"
    
    
    
    'Developer Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Developer"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Generate metadata from active workbook"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "GenerateSpreadsheetMetadataActiveWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Create spreadsheet from metadata"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CreateSpreadsheetFromMetadata"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Clean code"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CleanCode"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Import VBA modules from folder"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ImportVBAModulesFromFolder"
    
    'Other Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Other"
    
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Convert selected sheets to reporting sheets"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ConvertSelectedSheetsToReportingSheet"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Run looper"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "RunTableLooperOnActiveSheet"
    
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Export all tables in workbook as pipe delimited text"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ExportTablesInActiveWorkbookToPipeDelimtedText"
    
    
    Set MenuCategory = Nothing
    Set MenuSubcategory = Nothing
    Set MenuItem = Nothing
    
    
End Sub

