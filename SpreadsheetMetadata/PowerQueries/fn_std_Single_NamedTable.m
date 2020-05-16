(FolderPath as text, FileName as text, TableName as text)=>
let
    Source = Excel.Workbook(File.Contents(FolderPath & FileName), true, null),
    tbls = Table.SelectRows(Source, each [Kind] = "Table"),
    tbl = Table.SelectRows(tbls, each [Name] = TableName)[Data]{0}
in
    tbl