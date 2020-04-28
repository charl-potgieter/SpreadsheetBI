let    tbl = Table.FromRecords({        [Time Period = "MTD", Time Period Sort By Col = "1"],         [Time Period = "QTD", Time Period Sort By Col = "2"],         [Time Period = "YTD", Time Period Sort By Col = "3"],         [Time Period = "PY", Time Period Sort By Col = "4"]
        }),     ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"Time Period", type text},
            {"Time Period Sort By Col", type number}

        })in    ChangedType