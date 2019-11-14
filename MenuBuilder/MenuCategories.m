let
    Source = Csv.Document(File.Contents("C:\Users\charl\Dropbox\Dropbox_Charl\Computer_Technical\Programming_GitHub\SpreadsheetBI\MenuBuilder\MenuCategories.csv"),[Delimiter=",", Columns=2, Encoding=65001, QuoteStyle=QuoteStyle.None]),
    PromotedHeaders = Table.PromoteHeaders(Source, [PromoteAllScalars=true]),
    ChangedType = Table.TransformColumnTypes(PromotedHeaders,{{"Category", type text}, {"CategoryIndex", Int64.Type}})
in
    ChangedType