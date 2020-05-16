(FolderPath as text, FileName as text, SheetName as text)=>
let
    Source = Excel.Workbook(File.Contents(FolderPath & FileName), true, null),
    shts = Table.SelectRows(Source, each [Kind] = "Sheet"),
    tbl = Table.SelectRows(shts, each [Name] = SheetName)[Data]{0}
in
    tbl