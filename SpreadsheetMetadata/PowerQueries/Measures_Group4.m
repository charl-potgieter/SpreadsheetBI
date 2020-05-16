let    tbl = Table.FromRecords({        [NullHeader = ""]
        }),     ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"NullHeader", type text}

        })in    ChangedType