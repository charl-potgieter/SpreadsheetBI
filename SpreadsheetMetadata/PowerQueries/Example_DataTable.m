let    tbl = Table.FromRecords({        [Date = "31/01/2016", Description = "blah", SubCategory = "A", Amount = "1234"],         [Date = "28/02/2017", Description = "hello", SubCategory = "A", Amount = "100"],         [Date = "31/01/2016", Description = "blah", SubCategory = "A", Amount = "13334"],         [Date = "30/04/2018", Description = "hello", SubCategory = "B", Amount = "1550"],         [Date = "31/01/2016", Description = "zzzz", SubCategory = "A", Amount = "1034"],         [Date = "28/02/2017", Description = "hello", SubCategory = "A", Amount = "1500"],         [Date = "31/01/2016", Description = "zzzz", SubCategory = "A", Amount = "1734"],         [Date = "30/04/2018", Description = "hello", SubCategory = "B", Amount = "10"],         [Date = "28/08/2019", Description = "blah", SubCategory = "B", Amount = "1454"],         [Date = "31/05/2020", Description = "hello", SubCategory = "B", Amount = "1560"]
        }),     ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"Date", type date},
            {"Description", type text},
            {"SubCategory", type text},
            {"Amount", type number}

        })in    ChangedType