(OpeningBalance,AmortisationRatePerYear,StartDate)=>
let

    //Uncomment for debugging purposes
    //OpeningBalance = 500000,
    //AmortisationRatePerYear = 0.2,
    //StartDate = #date(2019,1,1),

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
    DeleteIndex