(FolderPath as text, FileName as text)=>
let
    Source = Csv.Document(File.Contents(FolderPath & FileName),[Delimiter="|", Encoding=1252, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true])
in
    PromotedHeaders