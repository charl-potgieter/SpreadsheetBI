(TextToConvert as text)=>
let
    MonthStart = #date(Number.From(Text.Start(TextToConvert, 4)), Number.From(Text.Range(TextToConvert, 4, 2)), 1),
    MonthEnd = Date.EndOfMonth(MonthStart)
in
    MonthEnd