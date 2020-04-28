let    tbl = Table.FromRecords({        [Description = "blah", Full Description = "This is blah"],         [Description = "hello", Full Description = "This is hello"],         [Description = "zzzz", Full Description = "This is zzzz"]
        }),     ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"Description", type text},
            {"Full Description", type text}

        })in    ChangedType