//(DataAccessMethod as text, optional fn_DataAccessCustom as function, SourceFolder as text, optional FilterFileNameFrom, optional FilterFileNameTo, optional IsDevMode as logical)=>
let
    

    /* 
        Consolidates files in SourceFolder with each file being read using fn_Single
        fn_Single needs to take 2 parameters, the file path and the file name
        The files are filtered based on file names using parameters FilterFileNameFrom and FilterFileNameTo
        These parameters need to be the same length and file names are truncated to this length for filtering purposes
    */


    
// ------------------------------------------------------------------------------------------------------------------
//                      Debugging
// ------------------------------------------------------------------------------------------------------------------

    //Uncomment for debugging    
    DataAccessMethod = "First sheet",
    fn_DataAccessCustom = ()=>none,
    SourceFolder = "D:\Onedrive\Documents_Charl\Computer_Technical\Programming_GitHub\SpreadsheetBI\Testing\Test_Consolidation\Test_Consolidation_Years",
    FilterFileNameFrom = "a",
    FilterFileNameTo = "b"
    IsDevMode = false,



// ------------------------------------------------------------------------------------------------------------------
//                      fn_DataAccessFirstSheet
// ------------------------------------------------------------------------------------------------------------------
    fn_DataAccessFirstSheet = 
    (Folder as text, FName as text)=>
    let
        Source = Excel.Workbook(File.Contents(Folder & FName), true, true),
        Navigation = Table.SelectRows(Source, each [Kind] = "Sheet")[Data]{0}
    in
        try Navigation otherwise error "Error in procedure fn_DataAccessFirstSheet",



// ------------------------------------------------------------------------------------------------------------------
//                      DataAccessFunctionSelected
// ------------------------------------------------------------------------------------------------------------------    
    DataAccessFunctionSelected = 
    let
        ReturnValue = if DataAccessMethod = "Custom" then
            fn_DataAccessCustom
        else if DataAccessMethod = "First sheet" then
            fn_DataAccessFirstSheet
        else
            error []             
    in
        try ReturnValue otherwise error "Error in procedure DataAccessFunctionSelected",



// ------------------------------------------------------------------------------------------------------------------
//                      fn_TypeConverter
// ------------------------------------------------------------------------------------------------------------------            
    fn_TypeConverter = 
    (TypeAsText as text)=>
    let
        
        ConverterRecord = [
            type null = type null,
            type logical = type logical,
            type number = type number,
            type time = type time,
            type date = type date,
            type datetime = type datetime,
            type datetimezone = type datetimezone,
            type duration = type duration,
            type text = type text,
            type binary = type binary,
            type type = type type,
            type list = type list,
            type record = type record,
            type table = type table,
            type function = type function,
            type anynonnull = type anynonnull,
            type none = type none,
            Int64.Type = Int64.Type,
            Currency.Type = Currency.Type,
            Percentage.Type = Percentage.Type
        ],
        
        ReturnValue = Record.Field(ConverterRecord, TypeAsText)

    in
        try ReturnValue otherwise error "Error in fn_TypeConverter",



// ------------------------------------------------------------------------------------------------------------------
//                      fn_GetRawData
// ------------------------------------------------------------------------------------------------------------------ 
    fn_GetRawData = 
    (SourceFolder as text, DataAccessFunction as function, FilterFileNameFrom, FilterFileNameTo)=>
    let
        // Get folder contents and filter out non-data files
        FolderContents = Folder.Files(SourceFolder),
        FilterOutNonData = Table.SelectRows(FolderContents, each
            Text.Upper([Name]) <> "README.TXT" and
            Text.Upper([Name]) <> "THUMBS.DB" and
            Text.Upper([Extension]) <> ".SQL" and
            Text.Start([Name], 1) <> "~"
            ),
        
        // Custom table type avoids types being lost on table expansion
        FirstTable = DataAccessFunction(FilterOutNonData[Folder Path]{0}, FilterOutNonData[Name]{0}),
        CustomTableType = Value.Type(FirstTable),
        AddTableCol = Table.AddColumn(FilterOutNonData, "tbl", each DataAccessFunction([Folder Path], [Name]), CustomTableType),
    
        // Filter data per parameters (using same number of characters)
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
        
        // If no file exists return an empty table to prevent an expand error
        Expand = if Table.RowCount(SelectTableCol) = 0 then
                #table({},{})
            else
                Table.ExpandTableColumn(
                    SelectTableCol, 
                    "tbl", 
                    Table.ColumnNames(SelectTableCol[tbl]{0}),
                    Table.ColumnNames(SelectTableCol[tbl]{0}))
                    
    in
        try Expand otherwise error "Error in procedure fn_GetRawData",
                
// ------------------------------------------------------------------------------------------------------------------
//                      fn_FilterParametersAreSameLength
// ------------------------------------------------------------------------------------------------------------------ 

    fn_FilterParametersAreSameLength = 
    (FilterFileNameFrom, FilerFileNameTo)=>
    let
        ReturnValue =if FilterFileNameFrom is null and FilterFileNameTo is null then
                true
            else if FilterFileNameFrom is null and not(FilterFileNameTo is null) then
                false
            else if not(FilterFileNameFrom is null) and FilterFileNameTo is null then
                false
            else if Text.Length(Text.From(FilterFileNameFrom)) <> Text.Length(Text.From(FilterFileNameTo)) then 
                false
            else
                true
    in
        try ReturnValue otherwise error "Eror in procedure fn_FilterParametersAreSameLength"


// ------------------------------------------------------------------------------------------------------------------
//                      Main
// ------------------------------------------------------------------------------------------------------------------ 


    ReturnValue = if fn_FilterParametersAreSameLength (FilterFileNameFrom, FilterFileNameTo) then
            "OK"
        else
            error "Filter parameters are different lengths"

in
    fn_GetRawData