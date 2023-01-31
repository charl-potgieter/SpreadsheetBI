(Folder, FName)=>
let
    Source = Excel.Workbook(File.Contents(Folder & FName), null, true),
    Sheet1_Sheet = Source{[Item="Sheet1",Kind="Sheet"]}[Data],
    RemovedTopRows = Table.Skip(Sheet1_Sheet,4),
    PromotedHeaders = Table.PromoteHeaders(RemovedTopRows, [PromoteAllScalars=true])
in
    PromotedHeaders