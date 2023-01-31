(Folder as text, optional SourceFileType as text, optional fn_CustomHeaderFunction as function)=>
let

/*
    Checks for consistency of file headers for all files in folder by generating the column headings by file
    SourceFileType paramater can be
     - Excel Data (data ion first tab of Excel sheet starting in row 1)
     - Excel Table (first table in first sheet of Excel file)
     - Csv
     - Other
    If Other is selected then fn_CustomHeaderFunction must be passed as a paramater taking Folder and
    Filename as its parameters and returning a list of column names
    This list can then be rolled up into Power BI or Power Pivot for reporting
*/

        
    fn_ColumnHeadersExcelData = 
    (Folder, FName)=>
    let
        Source = Excel.Workbook(File.Contents(Folder & FName), null, true),
        FirstSheet = Source[Data]{0},
        PromoteHeaders = Table.PromoteHeaders(FirstSheet, [PromoteAllScalars = true]), 
        ColumnNames = Table.ColumnNames(PromoteHeaders)
    in
        ColumnNames,
        

    fn_ColumnHeadersExcelTable = 
    (Folder, FName)=>
    let
        Source = Excel.Workbook(File.Contents(Folder & FName), null, true),
        FilteredOnTables = Table.SelectRows(Source, each ([Kind] = "Table")),
        FirstTable = FilteredOnTables[Data]{0},
        ColumnNames = Table.ColumnNames(FirstTable)
    in
        ColumnNames,


    fn_ColumnHeadersCSV = 
    (Folder, FName)=>
    let
        Source = Csv.Document(File.Contents(Folder & FName),[Delimiter=",", Encoding=1252, QuoteStyle=QuoteStyle.None]),
        PromoteHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
        ColumnNames = Table.ColumnNames(PromoteHeaders)
    in
        ColumnNames,

        
    // Function to get column headers depends on source file type as set in parameter
    fn_SelectedColumnHeaderFunction = if SourceFileType = "Excel Data" then
            fn_ColumnHeadersExcelData
        else if Text.Upper(SourceFileType) = "EXCEL DATA" then
            fn_ColumnHeadersExcelData
        else if Text.Upper(SourceFileType) = "EXCEL TABLE" then
            fn_ColumnHeadersExcelTable
        else if Text.Upper(SourceFileType) = "CSV" then
            fn_ColumnHeadersCSV
        else
            fn_CustomHeaderFunction,
        
    
    //Get folder contents and filter out non-data files
    FolderContents = Folder.Files(Folder),
    FilterOutNonData = Table.SelectRows(FolderContents, each
        Text.Upper([Name]) <> "README.TXT" and
        Text.Upper([Name]) <> "THUMBS.DB" and
        Text.Upper([Extension]) <> ".SQL" and
        Text.Start([Name], 1) <> "~"
        ),
        
        
    AddColumnNameList = Table.AddColumn(FilterOutNonData, "ColumnName", each fn_SelectedColumnHeaderFunction([Folder Path], [Name]), type list),
    SelectCols = Table.SelectColumns (AddColumnNameList, {"Folder Path", "Name", "ColumnName"}),
    Expand = Table.ExpandListColumn(SelectCols, "ColumnName"),
    ChangedType = Table.TransformColumnTypes(Expand,{{"Folder Path", type text}, {"Name", type text}, {"ColumnName", type text}})
    
in
    ChangedType