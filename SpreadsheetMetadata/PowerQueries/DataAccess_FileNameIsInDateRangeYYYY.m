(FileName as text, YearStart as number, YearEnd as number) =>
let
    //Checks wheter file name is inside date range where file name starts with YYYY
    YearFromFileName = Number.From(Text.Start(FileName, 4)),
    IsInRange = (YearFromFileName >= YearStart) and (YearFromFileName <= YearEnd)    
in
    IsInRange