(FileName as text, DateStart as date, DateEnd as date) =>
let
    //Checks wheter file name is inside date range where file name starts with YYYYMM
    YearFromFileName = Number.From(Text.Start(FileName, 4)),
    MonthFromFileName = Number.From(Text.Range(FileName, 4, 2)),
    MonthEndFromFileName = Date.EndOfMonth(#date(YearFromFileName, MonthFromFileName, 1)),
    IsInRange = (MonthEndFromFileName >= DateStart) and (MonthEndFromFileName <= DateEnd)    
in
    IsInRange