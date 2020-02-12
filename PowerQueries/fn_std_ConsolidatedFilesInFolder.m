let


    // -----------------------------------------------------------------------------------------------------------------------------------------
    //                      Documentation
    // -----------------------------------------------------------------------------------------------------------------------------------------
    
    Documentation_ = [
        Documentation.Name =  " fn_std_ConsolidatedFilsInFolder", 
        Documentation.Description = " Consolidates files in folder and filters by date." , 
        Documentation.LongDescription = " Consolidates files in folder and filters by date.", 
        Documentation.Category = " Table", 
        Documentation.Author = " Charl Potgieter"
        ],


    // -----------------------------------------------------------------------------------------------------------------------------------------
    //                      Function code
    // -----------------------------------------------------------------------------------------------------------------------------------------


    fn_=
    (
        FolderPath as text, 
        fn as function, 
        optional FileNamePrefix as text, 
        optional DateFrom as date, 
        optional DateTo as date, 
        optional DateFieldNameInUnderlyingFiles as text
    )=>
    let

        // Get folder contents and filter out Readme, .sql and temporary files starting with tildas
        FolderContents = Folder.Files(FolderPath),
        FilteredOutReadMeAndSQL = Table.SelectRows(FolderContents, each (Text.Upper([Name]) <> "README.TXT") and (Text.Upper([Extension]) <> ".SQL")),
        FilteredOutTildas = Table.SelectRows(FilteredOutReadMeAndSQL, each Text.Start([Name], 1) <> "~"),

        // Add period end field basd on FileNamePrefix
        AddPeriodEnd = if FileNamePrefix = "YYYY" then
            Table.AddColumn(FilteredOutTildas, "YearPerFileName", each Number.From(Text.Start([Name], 4)), type number)
        else if FileNamePrefix = "YYYYMM" then
            Table.AddColumn(FilteredOutTildas, "MonthEndPerFileName", each Date.EndOfMonth(#date(Number.From(Text.Start([Name], 4)), Number.From(Text.Range([Name], 4, 2)), 1)), type date)
        else
            FilteredOutTildas,

        // Filter on dates in file name if they exist adn DateFrom and DateTo are populated
        FilteredOnPeiodEndPerFileName = if FileNamePrefix = "YYYY" and DateFrom <> null and DateTo <> null then
            Table.SelectRows(AddPeriodEnd, each [YearPerFileName] >= Date.Year(DateFrom) and [YearPerFileName] <= Date.Year(DateTo))
        else if FileNamePrefix = "YYYYMM" then
            Table.SelectRows(AddPeriodEnd, each [MonthEndPerFileName] >= DateFrom and [MonthEndPerFileName] <= DateTo)
        else
            AddPeriodEnd,

        // Add single file tables, remove excess columns and expand
        AddTableCol = Table.AddColumn(FilteredOnPeiodEndPerFileName, "tbl", each fn([Folder Path], [Name]), Value.Type(fn(FilteredOutTildas[Folder Path]{0}, FilteredOutTildas[Name]{0}))),
        RemoveCols = Table.RemoveColumns(AddTableCol, {"Content", "Folder Path", "Name", "Extension", "Date accessed", "Date modified", "Date created", "Attributes"}),
        Expanded = Table.ExpandTableColumn(RemoveCols, "tbl", Table.ColumnNames(AddTableCol[tbl]{0})),

        // Filter on date in individual files if specified
        DateFiltered = if DateFieldNameInUnderlyingFiles <>  null then
            Table.SelectRows(Expanded, (_) => (Record.Field(_,
                DateFieldNameInUnderlyingFiles) >= DateFrom and Record.Field(_,
                DateFieldNameInUnderlyingFiles) <= DateTo))
            else
                Expanded

    in
        DateFiltered,


// -----------------------------------------------------------------------------------------------------------------------------------------
//                      Output
// -----------------------------------------------------------------------------------------------------------------------------------------

    type_ = type function (
        FolderPath as text, 
        fn as function, 
        optional FileNamePrefix as (type text meta [
                                Documentation.FieldCaption = "Select File Name Prefix",
                                Documentation.FieldDescription = "Utilised for filtering based on date",
                                Documentation.AllowedValues = {"YYYY", "YYYYMM"}
                                ]), 
        optional DateFrom as date, 
        optional DateTo as date, 
        optional DateFieldNameInUnderlyingFiles as text
        )
        as table meta Documentation_,

    // Replace the extisting type of the function with the individually defined
    Result =  Value.ReplaceType(fn_, type_)
 
 in 
    Result