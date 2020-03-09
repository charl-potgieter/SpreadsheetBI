let
    DateStart = Date.From(fn_std_Parameters("Date_Start")),
    DateEnd = Date.From(fn_std_Parameters("Date_End")),
    tbl = fn_std_DateTable(DateStart, DateEnd)
in
    tbl