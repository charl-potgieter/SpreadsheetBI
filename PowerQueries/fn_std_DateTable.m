let


    // -----------------------------------------------------------------------------------------------------------------------------------------
    //                      Documentation
    // -----------------------------------------------------------------------------------------------------------------------------------------
    
    Documentation_ = [
        Documentation.Name =  " fn_std_DateTable", 
        Documentation.Description = " Creates a standard date table." , 
        Documentation.LongDescription = " Creates a standard date table.", 
        Documentation.Category = " Table", 
        Documentation.Source = " tba", 
        Documentation.Author = " Charl Potgieter"
        ],


    // -----------------------------------------------------------------------------------------------------------------------------------------
    //                      Function code
    // -----------------------------------------------------------------------------------------------------------------------------------------

    fn_ = (DateFrom as date, DateTo as date)=>
    let
        
        // For debugging
        //DateFrom = #date(2016,5,4),
        //DateTo = #date(2017,8,15),

        // Get daylist
        YearStart = Date.StartOfYear(DateFrom),
        YearEnd = Date.EndOfYear(DateTo),
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
        InsertDayInWeek = Table.AddColumn(InsertCalendarQtr, "DayInWeek", each Date.DayOfWeek([Date]), Int64.Type)

    in
        InsertDayInWeek,


// -----------------------------------------------------------------------------------------------------------------------------------------
//                      Output
// -----------------------------------------------------------------------------------------------------------------------------------------

    type_ = type function (
        DateFrom as (type date),
        DateTo as (type date)
        )
        as table meta Documentation_,

    // Replace the extisting type of the function with the individually defined
    Result =  Value.ReplaceType(fn_, type_)
 
 in 
    Result