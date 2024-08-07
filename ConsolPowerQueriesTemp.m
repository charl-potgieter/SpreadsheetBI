[
    
Functions = 
[

/****************************************************************************************************************************************************
            Functions --> Data Access
****************************************************************************************************************************************************/

DataAccess = [

//+++++++++++++++++++++++++++++
//  ConsolidatedFilesInFolder
//+++++++++++++++++++++++++++++
ConsolidatedFilesInFolder = 
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


],

        

/****************************************************************************************************************************************************
            Functions --> Dates
****************************************************************************************************************************************************/
    
    
Dates = [
    
    
//++++++++++++++++++++++
//  DateTableStandard
//++++++++++++++++++++++


DateTableStandard = 
(FirstYear, LastYear)=>
let
    YearStart = #date(FirstYear, 1, 1),
    YearEnd = #date(LastYear, 12, 31),
    DayCount = Duration.Days(YearEnd - YearStart) +1,
    DayList =  List.Dates(YearStart, DayCount, #duration(1,0,0,0)),
    DayTable = Table.FromList(DayList, Splitter.SplitByNothing()),
    RenamedCols = Table.RenameColumns(DayTable, {"Column1", "Date"}),
    ChangedType = Table.TransformColumnTypes(RenamedCols, {{"Date", type date}}),

    //Insert year, qtr, month, day number
    InsertYear = Table.AddColumn(ChangedType, "Year", each Date.Year([Date]), Int64.Type),
    InsertQuarter = Table.AddColumn(InsertYear, "QuarterOfYear", each Date.QuarterOfYear([Date]), Int64.Type),
    InsertMonth = Table.AddColumn(InsertQuarter, "MonthOfYear", each Date.Month([Date]), Int64.Type),
    InsertDay = Table.AddColumn(InsertMonth, "DayOfMonth", each Date.Day([Date]), Int64.Type),

    //Insert end of Periods
    InsertEndOfYear = Table.AddColumn(InsertDay, "EndOfYear", each Date.EndOfYear([Date]), type date),
    InsertEndOfQtr = Table.AddColumn(InsertEndOfYear, "EndOfQtr", each Date.EndOfQuarter([Date]), type date),
    InsertEndOfMonth = Table.AddColumn(InsertEndOfQtr, "EndOfMonth", each Date.EndOfMonth([Date]), type date),
    InsertEndOfWeek = Table.AddColumn(InsertEndOfMonth, "EndOfWeek", each Date.EndOfWeek([Date]), type date),
    
    //Inset tests for end of periods
    InsertIsYearEnd = Table.AddColumn(InsertEndOfWeek, "IsEndOfYear", each [Date] = [EndOfYear], type logical),
    InsertIsQtrEnd = Table.AddColumn(InsertIsYearEnd, "IsEndOfQtr", each [Date] = [EndOfQtr], type logical),
    InsertIsMonthEnd = Table.AddColumn(InsertIsQtrEnd, "IsEndOfMonth", each [Date] = [EndOfMonth], type logical),
    InsertIsWeekEnd = Table.AddColumn(InsertIsMonthEnd, "IsEndOfWeek", each [Date] = [EndOfWeek], type logical),


    //Insert sundry fields
    InsertDateInt = Table.AddColumn(InsertIsWeekEnd, "DateInt", each ([Year] * 10000 + [MonthOfYear] * 100 + [DayOfMonth]), Int64.Type),
    InsertMonthName = Table.AddColumn(InsertDateInt, "MonthName", each Date.ToText([Date], "MMMM"), type text),
    InsertDayName = Table.AddColumn(InsertMonthName, "DayName", each Date.ToText([Date], "dddd"), type text),
    InsertCalendarMonth = Table.AddColumn(InsertDayName, "MonthInCalender", each (try(Text.Range([MonthName], 0, 3)) otherwise [MonthName]) & "-" & Text.End(Number.ToText([Year]), 2), type text),
    InsertCalendarQtr = Table.AddColumn(InsertCalendarMonth, "QuarterInCalendar", each "Q" & Number.ToText([QuarterOfYear]) &" " & Number.ToText([Year]), type text),
    InsertDayInWeek = Table.AddColumn(InsertCalendarQtr, "DayInWeek", each Date.DayOfWeek([Date]), Int64.Type),
    AddDaysInYearCol = Table.AddColumn(InsertDayInWeek, "DaysInYear",each Date.DayOfYear(Date.EndOfYear([Date])), Int64.Type)
    
in
    AddDaysInYearCol,



//+++++++++++++++++++++++++++++++++
//  DateTableJuneYearEnd
//+++++++++++++++++++++++++++++++++

DateTableJuneYearEnd  = 
(FirstYear, LastYear)=>
let

    // Get daylist
    YearStart = #date(FirstYear-1,7,1),
    YearEnd = #date(LastYear, 6, 30),
    DayCount = Duration.Days(YearEnd - YearStart) +1,
    DayList =  List.Dates(YearStart, DayCount, #duration(1,0,0,0)),
    DayTable = Table.FromList(DayList, Splitter.SplitByNothing()),
    RenamedCols = Table.RenameColumns(DayTable, {"Column1", "Date"}),
    ChangedType = Table.TransformColumnTypes(RenamedCols, {{"Date", type date}}),

    //Convert calender qtr numbers to June year end quarter numbers
    fn_ConvertQtrsToJuneYearEnds = (x)=>if x = 1 then 3
            else if x = 2 then 4
            else if x = 3 then 1
            else 2,

    InsertQuarter = Table.AddColumn(ChangedType, "QuarterOfYear", each fn_ConvertQtrsToJuneYearEnds(Date.QuarterOfYear([Date])), Int64.Type),

    //Convert calender month numbers to June year end month numbers
    fn_ConvertMonthToJuneYearEnds = (x)=>if x <=6 then x+6 else x-6,        

    InsertMonth = Table.AddColumn(InsertQuarter, "MonthOfYear", each fn_ConvertMonthToJuneYearEnds(Date.Month([Date])), Int64.Type),
    InsertDay = Table.AddColumn(InsertMonth, "DayOfMonth", each Date.Day([Date]), Int64.Type),

    //Insert end of Periods
    InsertEndOfYear = Table.AddColumn(InsertDay, "EndOfYear", each 
            if Date.Month([Date])<=6 then 
                    #date(Date.Year([Date]), 6, 30)
            else
                    #date(Date.Year([Date]) +1, 6, 30)
            , type date),

    InsertEndOfQtr = Table.AddColumn(InsertEndOfYear, "EndOfQtr", each Date.EndOfQuarter([Date]), type date),
    InsertEndOfMonth = Table.AddColumn(InsertEndOfQtr, "EndOfMonth", each Date.EndOfMonth([Date]), type date),
    InsertEndOfWeek = Table.AddColumn(InsertEndOfMonth, "EndOfWeek", each Date.EndOfWeek([Date]), type date),
    
    //Inset tests for end of periods
    InsertIsYearEnd = Table.AddColumn(InsertEndOfWeek, "IsEndOfYear", each [Date] = [EndOfYear], type logical),
    InsertIsQtrEnd = Table.AddColumn(InsertIsYearEnd, "IsEndOfQtr", each [Date] = [EndOfQtr], type logical),
    InsertIsMonthEnd = Table.AddColumn(InsertIsQtrEnd, "IsEndOfMonth", each [Date] = [EndOfMonth], type logical),
    InsertIsWeekEnd = Table.AddColumn(InsertIsMonthEnd, "IsEndOfWeek", each [Date] = [EndOfWeek], type logical),


    //Insert sundry fields
    InsertDateInt = Table.AddColumn(InsertIsWeekEnd, "DateInt", each (Date.Year([Date]) * 10000 + Date.Day([Date]) * 100 + Date.Day([Date])), Int64.Type),
    InsertMonthName = Table.AddColumn(InsertDateInt, "MonthName", each Date.ToText([Date], "MMMM"), type text),
    InsertDayName = Table.AddColumn(InsertMonthName, "DayName", each Date.ToText([Date], "dddd"), type text),
    InsertCalendarMonth = Table.AddColumn(InsertDayName, "MonthInCalender", each (try(Text.Range([MonthName], 0, 3)) otherwise [MonthName]) & "-" & Text.End(Number.ToText(Date.Year([Date])), 2), type text),
    InsertCalendarQtr = Table.AddColumn(InsertCalendarMonth, "QuarterInCalendar", each "Q" & Number.ToText([QuarterOfYear]), type text),
    InsertDayInWeek = Table.AddColumn(InsertCalendarQtr, "DayInWeek", each Date.DayOfWeek([Date]), Int64.Type),


    //Function days in tax year
    fn_DaysInTaxYear=
    (ye)=>
    let
            FirstDayOfYear = #date(Date.Year(ye)-1, 7,1),
            NumDates = Duration.Days(ye - FirstDayOfYear) +1
    in
            NumDates,

    AddDaysInYearCol = Table.AddColumn(InsertDayInWeek, "DaysInYear",each fn_DaysInTaxYear([EndOfYear]), Int64.Type)
    
in
    AddDaysInYearCol,


//+++++++++++++++++++++++++++++++++
//  fn_DatesBetween
//+++++++++++++++++++++++++++++++++
fn_DatesBetween = 

let 
    // Credit for below code = Imke Feldman Imke Feldmann: www.TheBIccountant.com

    // ----------------------- Documentation ----------------------- 

    documentation_ = [
        Documentation.Name =  " Dates.DatesBetween", 
        Documentation.Description = " Creates a list of dates according to the chosen interval between Start and End. Allowed values for 3rd parameter: ""Year"", ""Quarter"", ""Month"", ""Week"" or ""Day""." , 
        Documentation.LongDescription = " Creates a list of dates according to the chosen interval between Start and End. The dates created will always be at the end of the interval, so could be in the future if today is chosen.", 
        Documentation.Category = " Table", 
        Documentation.Source = " http://www.thebiccountant.com/2017/12/11/date-datesbetween-retrieve-dates-between-2-dates-power-bi-power-query/ . ", 
        Documentation.Author = " Imke Feldmann: www.TheBIccountant.com . ", 
        Documentation.Examples = {[Description =  " Check this blogpost: http://www.thebiccountant.com/2017/12/11/date-datesbetween-retrieve-dates-between-2-dates-power-bi-power-query/ ." , 
            Code = "", 
            Result = ""]}
        ],

    // ----------------------- Function Code ----------------------- 
    
    function_ =  (From as date, To as date, optional Selection as text ) =>
    let

        // Create default-value "Day" if no selection for the 3rd parameter has been made
        TimeInterval = if Selection = null then "Day" else Selection,

        // Table with different values for each case
        CaseFunctions = #table({"Case", "LastDateInTI", "TypeOfAddedTI", "NumberOfAddedTIs"},
                {   {"Day", Date.From, Date.AddDays, Number.From(To-From)+1},
                    {"Week", Date.EndOfWeek, Date.AddWeeks, Number.RoundUp((Number.From(To-From)+1)/7)},
                    {"Month", Date.EndOfMonth, Date.AddMonths, (Date.Year(To)*12+Date.Month(To))-(Date.Year(From)*12+Date.Month(From))+1},
                    {"Quarter", Date.EndOfQuarter, Date.AddQuarters, (Date.Year(To)*4+Date.QuarterOfYear(To))-(Date.Year(From)*4+Date.QuarterOfYear(From))+1},
                    {"Year", Date.EndOfYear, Date.AddYears,Date.Year(To)-Date.Year(From)+1} 
                } ),

        // Filter table on selected case
        Case = CaseFunctions{[Case = TimeInterval]},
        
        // Create list with dates: List with number of date intervals -> Add number of intervals to From-parameter -> shift dates at the end of each respective interval	
        DateFunction = List.Transform({0..Case[NumberOfAddedTIs]-1}, each Function.Invoke(Case[LastDateInTI], {Function.Invoke(Case[TypeOfAddedTI], {From, _})}))
    in
        DateFunction,

    // ----------------------- New Function Type ----------------------- 

    type_ = type function (
        From as (type date),
        To as (type date),
        optional Selection as (type text meta [
                                Documentation.FieldCaption = "Select Date Interval",
                                Documentation.FieldDescription = "Select Date Interval, if nothing selected, the default value will be ""Day""",
                                Documentation.AllowedValues = {"Day", "Week", "Month", "Quarter", "Year"}
                                ])
            )
        as table meta documentation_,

    // Replace the extisting type of the function with the individually defined
    Result =  Value.ReplaceType(function_, type_)
 
 in 

    Result,

//+++++++++++++++++++++++++++++++++
//  fn_FileNameIsInDateRangeYYYY
//+++++++++++++++++++++++++++++++++    
FileNameIsInDateRangeYYYY =     
(FileName as text, YearStart as number, YearEnd as number) =>
let
    //Checks wheter file name is inside date range where file name starts with YYYY
    YearFromFileName = Number.From(Text.Start(FileName, 4)),
    IsInRange = (YearFromFileName >= YearStart) and (YearFromFileName <= YearEnd)    
in
    IsInRange,


//+++++++++++++++++++++++++++++++++
//  fn_FileNameIsInDateRangeYYYY
//+++++++++++++++++++++++++++++++++    
FileNameIsInDateRangeYYYYMM = 
(FileName as text, DateStart as date, DateEnd as date) =>
let
    //Checks wheter file name is inside date range where file name starts with YYYYMM
    YearFromFileName = Number.From(Text.Start(FileName, 4)),
    MonthFromFileName = Number.From(Text.Range(FileName, 4, 2)),
    MonthEndFromFileName = Date.EndOfMonth(#date(YearFromFileName, MonthFromFileName, 1)),
    IsInRange = (MonthEndFromFileName >= DateStart) and (MonthEndFromFileName <= DateEnd)    
in
    IsInRange



], 


/****************************************************************************************************************************************************
            Functions --> Other
****************************************************************************************************************************************************/

Other = [

//+++++++++++++++++++++++++++++++++
//  fn_ConvertAllColumnsToText
//+++++++++++++++++++++++++++++++++
fn_ConvertAllColumnsToText = 
(tbl)=>
let
    ConversionList = List.Transform(Table.ColumnNames(tbl), each {_, type text}),
    Converted = Table.TransformColumnTypes(tbl, ConversionList)
in
    Converted,


//+++++++++++++++++++++++++++++++++
//  fn_StraightLineAmortisation
//+++++++++++++++++++++++++++++++++        
fn_StraightLineAmortisation =         
(OpeningBalance,AmortisationRatePerYear,StartDate)=>
let

    //Uncomment for debugging purposes
    OpeningBalance = 500000,
    AmortisationRatePerYear = 0.2,
    StartDate = #date(2019,1,1),

    NumberOfMonths = (1 / AmortisationRatePerYear) * 12,

    IndexList = {1..NumberOfMonths},
    ConvertToTable = Table.FromList(IndexList, Splitter.SplitByNothing(), {"Index"}),
    ChangedIndexType = Table.TransformColumnTypes(ConvertToTable,{{"Index", Int64.Type}}),
    AddEndOfMonth = Table.AddColumn(ChangedIndexType, "End Of Month", each Date.EndOfMonth(Date.AddMonths(StartDate, [Index]-1)), type date),
    AddOpeningBalance = Table.AddColumn(AddEndOfMonth, "Opening Balance", each (NumberOfMonths - ([Index]-1)) / NumberOfMonths * OpeningBalance, type number),
    AddAmortisation = Table.AddColumn(AddOpeningBalance, "Amortisation", each OpeningBalance / NumberOfMonths, type number),
    AddClosingBalance = Table.AddColumn(AddAmortisation, "Closing Balance", each [Opening Balance] - [Amortisation], type number),
    DeleteIndex = Table.RemoveColumns(AddClosingBalance,{"Index"})
in
    DeleteIndex,
    

//+++++++++++++++++++++++++++++++++
//  fn_Parameters
//+++++++++++++++++++++++++++++++++            
fn_Parameters = 
let
    
    Documentation_ = [
        Documentation.Name =  " fn_std_Parameters", 
        Documentation.Description = " Returns parameter value set out in  tbl_Parameters" , 
        Documentation.LongDescription = "  Returns parameter value set out in  tbl_Parameters", 
        Documentation.Category = "Text",  
        Documentation.Author = " Charl Potgieter"
        ],

    fn_=
    (parameter as text)=>
    let
        Source = Excel.CurrentWorkbook(){[Name = "tbl_Parameters"]}[Content],
        FilteredRows = Table.SelectRows(Source, each [Parameter] = parameter),
        ReturnValue = FilteredRows[Value]{0}
    in
        ReturnValue,

    type_ = type function (
        parameter as (type text)
        )
        as text meta Documentation_,

    // Replace the extisting type of the function with the individually defined
    Result =  Value.ReplaceType(fn_, type_)
 
 in 
    Result


        
        
]  // End Functions other



],  // End Functions


Tables = 
[

/****************************************************************************************************************************************************
            Tables --> Sample data
****************************************************************************************************************************************************/



Sample data = [

//+++++++++++++++++++++++++++++++++
//  SampleDataTable
//+++++++++++++++++++++++++++++++++
DataTable = 
let

    tbl = Table.FromRecords({
        [Date = "42400", Foreign Key = "blah", SubCategory = "A", Amount = "1234"], 
        [Date = "42794", Foreign Key = "hello", SubCategory = "A", Amount = "100"], 
        [Date = "42400", Foreign Key = "blah b", SubCategory = "A", Amount = "13334"], 
        [Date = "43220", Foreign Key = "hello", SubCategory = "B", Amount = "1550"], 
        [Date = "42400", Foreign Key = "zzzz", SubCategory = "A", Amount = "1034"], 
        [Date = "42794", Foreign Key = "hello", SubCategory = "A", Amount = "1500"], 
        [Date = "42400", Foreign Key = "zzzz", SubCategory = "A", Amount = "1734"], 
        [Date = "43220", Foreign Key = "hello b", SubCategory = "B", Amount = "10"], 
        [Date = "43705", Foreign Key = "blah", SubCategory = "B", Amount = "1454"], 
        [Date = "43982", Foreign Key = "hello", SubCategory = "B", Amount = "1560"]
        }), 

    ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"Date", type text},
            {"Foreign Key", type text},
            {"SubCategory", type text},
            {"Amount", type number}

        })

in
    ChangedType,



//+++++++++++++++++++++++++++++++++
//  SampleLookupTable
//+++++++++++++++++++++++++++++++++
LookupTable = 
let

    tbl = Table.FromRecords({
        [Primary Key = "blah", Full Description = "This is blah"], 
        [Primary Key = "hello", Full Description = "This is hello"], 
        [Primary Key = "zzzz", Full Description = "This is zzzz"]
        }), 

    ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"Primary Key", type text},
            {"Full Description", type text}

        })

in
    ChangedType




],
 
/****************************************************************************************************************************************************
            Tables --> Other
****************************************************************************************************************************************************/
 
 
Other = [


//+++++++++++++++++++++++++++++++++
//  NumberScale
//+++++++++++++++++++++++++++++++++
NumberScale = 
let

    tbl = Table.FromRecords({
        [ShowValuesAs = "CCY", DivideBy = "1"], 
        [ShowValuesAs = "'000", DivideBy = "1000"], 
        [ShowValuesAs = "m", DivideBy = "1000000"]
        }), 

    ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"ShowValuesAs", type text},
            {"DivideBy", type number}

        })

in
    ChangedType,



//+++++++++++++++++++++++++++++++++
//  TimePeriods
//+++++++++++++++++++++++++++++++++
TimePeriods = 
let

    tbl = Table.FromRecords({
        [Time Period = "MTD", Time Period Sort By Col = "1"], 
        [Time Period = "QTD", Time Period Sort By Col = "2"], 
        [Time Period = "YTD", Time Period Sort By Col = "3"], 
        [Time Period = "PY", Time Period Sort By Col = "4"], 
        [Time Period = "Total", Time Period Sort By Col = "5"], 
        [Time Period = "As at date", Time Period Sort By Col = "6"],
        [Time Period = "As at month end", Time Period Sort By Col = "7"]
        }), 

    ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"Time Period", type text},
            {"Time Period Sort By Col", Int64.Type}

        })

in
    ChangedType



]
            
]


]