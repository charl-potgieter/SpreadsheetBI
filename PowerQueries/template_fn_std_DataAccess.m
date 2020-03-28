//Uncomment parameter once debugging complete
//(LoadData as logical)=>
let

    //Delete this once parameter is uncommented.
    LoadData  = true,

    DateStart = Date.From(fn_std_Parameters("Date_Start")),
    DateEnd = Date.From(fn_std_Parameters("Date_End")),

    YearStart = Date.Year(Date.From(fn_std_Parameters("Date_Start"))),
    YearEnd = Date.Year(Date.From(fn_std_Parameters("Date_End"))),

    FolderPath = "XXXXXX"
    

    // **** Uncomment one of the below options and change the last line of file  to read tblRaw***

    //No filter
    //tblRaw = fn_std_ConsolidatedFilesInFolder(FolderParth, fn_Single_XXXXX, LoadData, null, null, null, XXX_optional_sheet_or_table_name)

    //Filter files on year name
    // tblRaw = fn_std_ConsolidatedFilesInFolder(FolderParth, fn_Single_XXXXX, LoadData, fn_std_FileNameIsInDateRangeYYYY, YearStart, YearEnd, XXX_optional_sheet_or_table_name)

    //Filter files on month name 
    //tblRaw = fn_std_ConsolidatedFilesInFolder(FolderParth, fn_Single_XXXXX, LoadData, fn_std_FileNameIsInDateRangeYYYYMM, DateStart, DateEnd, XXX_optional_sheet_or_table_name)

in
    FolderPath