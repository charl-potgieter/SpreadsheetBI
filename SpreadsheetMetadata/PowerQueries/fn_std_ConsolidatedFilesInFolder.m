(
    FolderPath as text, 
    fn_SingleFile as function, 
    LoadData as logical,
    optional fn_FilterBasedOnFileName as function,
    optional FilterFromValue,
    optional FilterToValue,
    optional Additional_fn_SingleFileParameter as text      // utilsed for example to specify specific sheet name or table in fn_SingleFile
)=>
let

    // Get folder contents and filter out Readme, .sql and temporary files starting with tildas
    FolderContents = Folder.Files(FolderPath),
    FilteredOutReadMeAndSQL = Table.SelectRows(FolderContents, each (Text.Upper([Name]) <> "README.TXT") and (Text.Upper([Extension]) <> ".SQL")),
    FilteredOutTildas = Table.SelectRows(FilteredOutReadMeAndSQL, each Text.Start([Name], 1) <> "~"),

    //Restrict to one file if no data load
    ReturnOnlyIfLoadRequested = if LoadData then FilteredOutTildas else Table.FirstN(FilteredOutTildas, 1),

    //Filter Files based on filter function
    FilteredFile = if (fn_FilterBasedOnFileName <> null and Table.RowCount(ReturnOnlyIfLoadRequested) > 1 ) then
        Table.SelectRows(ReturnOnlyIfLoadRequested, each fn_FilterBasedOnFileName([Name], FilterFromValue, FilterToValue))
    else
        ReturnOnlyIfLoadRequested,

    // Add single file tables, remove excess columns and expand
    AddTableCol = if Additional_fn_SingleFileParameter = null then
            Table.AddColumn(FilteredFile, "tbl", each fn_SingleFile([Folder Path], [Name]))
        else
            Table.AddColumn(FilteredFile, "tbl", each fn_SingleFile([Folder Path], [Name], Additional_fn_SingleFileParameter)),
    RemoveCols = Table.RemoveColumns(AddTableCol, {"Content", "Extension", "Date accessed", "Date modified", "Date created", "Attributes"}),
    Expanded = Table.ExpandTableColumn(RemoveCols, "tbl", Table.ColumnNames(AddTableCol[tbl]{0})),

    // Filter at a data row level if required
    ReturnOnlyOneDataRowIfRequired = if LoadData then Expanded else Table.FirstN(Expanded, 1)

in
    ReturnOnlyOneDataRowIfRequired