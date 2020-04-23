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
    MenuItem.Caption = "Index page"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "InsertIndexPageActiveWorkbook"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Formatted sheet"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "InsertFormattedSheetIntoActiveWorkbook"
    
    
    
        
    'Pivot Table Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Pivot table"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Flatten"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "FormatPivotTableFlatten"
    
    
    'Data Model Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Data model"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Create BI spreadsheet"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "CreateBiSpreadsheet"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Write measures, columns and relationships to sheets"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "WritesMeasuresColumnsRelationshipsToSheetsEntryPoint"
    
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
    MenuItem.Caption = "Import power queries in folder non-recursive"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ImportPowerQueriesFromSelectedFolderRecursive"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Import selected power queries"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "ImportSelectedPowerQueries"
    
    'Other Menu
    Set MenuCategory = cb.Controls.Add(Type:=msoControlPopup)
    MenuCategory.Caption = "Other"
    
    Set MenuItem = MenuCategory.Controls.Add(Type:=msoControlButton)
    MenuItem.Caption = "Run looper"
    MenuItem.OnAction = "'" & ThisWorkbook.Name & "'!" & "TableLooper"
    
    
    
End Sub

