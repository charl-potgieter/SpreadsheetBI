let

    tbl = Table.FromRecords({
        [Date = "42400", Foreign Key = "blah", SubCategory = "A", Amount = "1234"], 
        [Date = "42794", Foreign Key = "hello", SubCategory = "A", Amount = "100"], 
        [Date = "42400", Foreign Key = "blah b", SubCategory = "A", Amount = "13334"], 
        [Date = "43220", Foreign Key = "hello", SubCategory = "B", Amount = "1550"], 
        [Date = "42400", Foreign Key = "zzzz", SubCategory = "A", Amount = "1034"], 
        [Date = "42794", Foreign Key = "hello", SubCategory = "A", Amount = "1500"], 
        [Date = "42400", Foreign Key = "zzzz", SubCategory = "A", Amount = "1734"], 
        [Date = "43220", Foreign Key = "hello b", SubCategory = "B", Amount = "10"], 
        [Date = "43705", Foreign Key = "blah", SubCategory = "B", Amount = "1454"], 
        [Date = "43982", Foreign Key = "hello", SubCategory = "B", Amount = "1560"]
        }), 

    ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"Date", type text},
            {"Foreign Key", type text},
            {"SubCategory", type text},
            {"Amount", type number}

        })

in
    ChangedType