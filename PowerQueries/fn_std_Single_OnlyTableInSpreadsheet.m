(FolderPath as text, FileName as text) =>
let
    Source = Excel.Workbook(File.Contents(FolderPath & FileName), null, true),
    tbl = Table.SelectRows(Source, each [Kind] = "Table")[Data]{0}
in
    tbl