let

    tbl = Table.FromRecords({
        [ShowValuesAs = "CCY", DivideBy = "1"], 
        [ShowValuesAs = "'000", DivideBy = "1000"], 
        [ShowValuesAs = "m", DivideBy = "1000000"]
        }), 

    ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"ShowValuesAs", type text},
            {"DivideBy", type number}

        })

in
    ChangedType