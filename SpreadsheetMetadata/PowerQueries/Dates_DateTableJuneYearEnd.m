(FirstYear, LastYear)=>
let

   // FirstYear = 2010,
   // LastYear = 2012,

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
    InsertCalendarYear = Table.AddColumn(InsertDay, "CalendarYear", each Date.Year([Date]), Int64.Type),

    //Insert end of Periods
    InsertEndOfYearJune = Table.AddColumn(InsertCalendarYear, "EndOfYear_June", each 
            if Date.Month([Date])<=6 then 
                    #date(Date.Year([Date]), 6, 30)
            else
                    #date(Date.Year([Date]) +1, 6, 30)
            , type date),


    InsertEndOfYearJuneText = Table.AddColumn(InsertEndOfYearJune, "EndOfYear_JuneText", each "YE June " & Text.From(Date.Year([EndOfYear_June])), type text),

    InsertEndOfYearCalendar = Table.AddColumn(InsertEndOfYearJuneText, "EndOfYear_Calendar", each Date.EndOfYear([Date]), type date),

    InsertEndOfQtr = Table.AddColumn(InsertEndOfYearCalendar, "EndOfQtr", each Date.EndOfQuarter([Date]), type date),
    InsertEndOfMonth = Table.AddColumn(InsertEndOfQtr, "EndOfMonth", each Date.EndOfMonth([Date]), type date),
    InsertEndOfWeek = Table.AddColumn(InsertEndOfMonth, "EndOfWeek", each Date.EndOfWeek([Date]), type date),
    
    //Inset tests for end of periods
    InsertIsYearEnd_June = Table.AddColumn(InsertEndOfWeek, "IsEndOfJuneYear_June", each [Date] = [EndOfYear_June], type logical),
    InsertIsYearEnd_Calendar = Table.AddColumn(InsertIsYearEnd_June, "IsEndOfYear_Calendar", each [Date] = [EndOfYear_Calendar], type logical),
    InsertIsQtrEnd = Table.AddColumn(InsertIsYearEnd_Calendar, "IsEndOfQtr", each [Date] = [EndOfQtr], type logical),
    InsertIsMonthEnd = Table.AddColumn(InsertIsQtrEnd, "IsEndOfMonth", each [Date] = [EndOfMonth], type logical),
    InsertIsWeekEnd = Table.AddColumn(InsertIsMonthEnd, "IsEndOfWeek", each [Date] = [EndOfWeek], type logical),


    //Insert sundry fields
    InsertDateInt = Table.AddColumn(InsertIsWeekEnd, "DateInt", each (Date.Year([Date]) * 10000 + Date.Day([Date]) * 100 + Date.Day([Date])), Int64.Type),
    InsertMonthName = Table.AddColumn(InsertDateInt, "MonthName", each Date.ToText([Date], "MMMM"), type text),
    InsertDayName = Table.AddColumn(InsertMonthName, "DayName", each Date.ToText([Date], "dddd"), type text),
    InsertMonthYearText = Table.AddColumn(InsertDayName, "MonthYearText", each (try(Text.Range([MonthName], 0, 3)) otherwise [MonthName]) & "-" & Text.End(Number.ToText(Date.Year([Date])), 2), type text),
    InsertQtrYearText = Table.AddColumn(InsertMonthYearText, "QuarterYearText", each "Q" & Number.ToText([QuarterOfYear]), type text),
    InsertDayInWeek = Table.AddColumn(InsertQtrYearText, "DayInWeek", each Date.DayOfWeek([Date]), Int64.Type),


    //Function days in tax year
    fn_DaysInTaxYear=
    (ye)=>
    let
            FirstDayOfYear = #date(Date.Year(ye)-1, 7,1),
            NumDates = Duration.Days(ye - FirstDayOfYear) +1
    in
            NumDates,

    AddDaysInJuneYearCol = Table.AddColumn(InsertDayInWeek, "DaysInTaxYear",each fn_DaysInTaxYear([EndOfYear_June]), Int64.Type),
    AddDaysInCalendarYear = Table.AddColumn(AddDaysInJuneYearCol, "DaysInCalendarYear", each Date.DayOfYear([EndOfYear_Calendar]), Int64.Type)
in
    AddDaysInCalendarYear