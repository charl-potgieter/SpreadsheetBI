let
    Source = Excel.CurrentWorkbook(){[Name="Table5"]}[Content],
    RemovedCols = Table.RemoveColumns(Source,{"Balance Sheet Allocation"}),
    ChangedType = Table.TransformColumnTypes(RemovedCols,{{"Account Number", type text}, {"Account Name", type text}, {"1", type number}, {"2", type number}, {"3", type number}, {"4", type number}, {"5", type number}, {"6", type number}, {"7", type number}, {"8", type number}, {"9", type number}, {"10", type number}, {"11", type number}, {"12", type number}, {"13", type number}, {"14", type number}, {"15", type number}, {"16", type number}, {"17", type number}, {"18", type number}, {"19", type number}, {"20", type number}, {"21", type number}, {"22", type number}, {"23", type number}, {"24", type number}}),
    UnpivotedCols = Table.UnpivotOtherColumns(ChangedType, {"Account Number", "Account Name"}, "MonthCount", "Amount"),
    ChangedType2 = Table.TransformColumnTypes(UnpivotedCols,{{"MonthCount", Int64.Type}}),
    AddMonthOfYear = Table.AddColumn(ChangedType2, "MonthOfYear", each if [MonthCount]>12 then [MonthCount]-12 else [MonthCount], Int64.Type),
    AddYearColumn = Table.AddColumn(AddMonthOfYear, "Year", each if Number.From([MonthCount])<=12 then 2018 else 2019, Int64.Type),
    AddEndOfMonthCol = Table.AddColumn(AddYearColumn, "End Of Month", each Date.EndOfMonth(#date([Year], [MonthOfYear], 1)), type date),
    SelectCols = Table.SelectColumns(AddEndOfMonthCol,{"End Of Month", "Account Number", "Amount" })
in
    SelectCols