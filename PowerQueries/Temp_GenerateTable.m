let

    tbl = Table.FromRecords({
        [Col1 = "a", Col2 = 3], 
        [Col1 = "b", Col2=5]}),

    ChangedType = Table.TransformColumnTypes(
        tbl,
        {{"Col1", type text}, 
        {"Col2", Int64.Type}})
        
in
    ChangedType