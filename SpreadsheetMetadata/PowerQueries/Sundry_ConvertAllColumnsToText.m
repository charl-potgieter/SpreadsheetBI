(tbl)=>
let
    ConversionList = List.Transform(Table.ColumnNames(tbl), each {_, type text}),
    Converted = Table.TransformColumnTypes(tbl, ConversionList)
in
    Converted