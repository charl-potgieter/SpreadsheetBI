(fn_Single as function, SourceFolder as text, optional FilterFileNameFrom, optional FilterFileNameTo, optional IsDevMode as logical)=>
let
    
    /* 
        Consolidates files in SourceFolder with each file being read using fn_Single
        fn_Single needs to take 2 parameters, the file path and the file name
        The files are filtered based on file names using parameters FilterFileNameFrom and FilterFileNameTo
        These parameters need to be the same length and file names are truncated to this length for filtering purposes
    */


    //Get folder contents and filter out non-data files
    FolderContents = Folder.Files(SourceFolder),
    FilterOutNonData = Table.SelectRows(FolderContents, each
        Text.Upper([Name]) <> "README.TXT" and
        Text.Upper([Name]) <> "THUMBS.DB" and
        Text.Upper([Extension]) <> ".SQL" and
        Text.Start([Name], 1) <> "~"
        ),
        
    //Custom table type avoids types being lost on table expansion
    FirstTable = fn_Single(FilterOutNonData[Folder Path]{0}, FilterOutNonData[Name]{0}),
    CustomTableType = Value.Type(FirstTable),
    AddTableCol = Table.AddColumn(FilterOutNonData, "tbl", each fn_Single([Folder Path], [Name]), CustomTableType),
    
    //Filter data per parameters (using same number of characters)
    FilterFileNameFromText = Text.From(FilterFileNameFrom),
    FilterFileNameToText = Text.From(FilterFileNameTo),
    FilterCharacterLength = Text.Length(FilterFileNameFromText),
    AddFilterCol = Table.AddColumn(AddTableCol, "FilterCol", each Text.Start([Name], FilterCharacterLength), type text),
    FilterFiles = Table.SelectRows(AddFilterCol, each ([FilterCol] >= FilterFileNameFromText) and ([FilterCol] <= FilterFileNameToText)), 
    DevMode_FilterOneFile = if IsDevMode is null then
            FilterFiles
        else if IsDevMode then 
            Table.FirstN(FilterFiles, 1) 
        else 
            FilterFiles,
    
    SelectTableCol = Table.SelectColumns(DevMode_FilterOneFile, {"tbl"}),
    
    //If no file exists return an empty table to prevent an expand error
    Expand = if Table.RowCount(SelectTableCol) = 0 then
            #table({},{})
        else
            Table.ExpandTableColumn(
                SelectTableCol, 
                "tbl", 
                Table.ColumnNames(SelectTableCol[tbl]{0}),
                Table.ColumnNames(SelectTableCol[tbl]{0})),
                
                
    CheckForMismatchParameterLenghth = if Text.Length(Text.From(FilterFileNameFrom)) <> Text.Length(Text.From(FilterFileNameTo)) then 
            error [
                Reason = "Business Rule Violated", 
                Message = "Item codes must start with a letter", 
                Detail = "Non-conforming Item Code: 456"
            ]
        else
            Expand

in
    CheckForMismatchParameterLenghth