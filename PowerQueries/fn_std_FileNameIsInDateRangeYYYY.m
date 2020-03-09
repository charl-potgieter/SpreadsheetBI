/*---------------------------------------------------------------------------------
    Checks wheter file name is inside date range where file name starts with YYYY
---------------------------------------------------------------------------------*/

(FileName as text, DateStart as date, DateEnd as date) =>
let
    YearStart = Date.Year(DateStart),
    YearEnd = Date.Year(DateEnd),
    YearFromFileName = Number.From(Text.Start(FileName, 4)),
    IsInRange = (YearFromFileName >= YearStart) and (YearFromFileName <= YearEnd)    
in
    IsInRange