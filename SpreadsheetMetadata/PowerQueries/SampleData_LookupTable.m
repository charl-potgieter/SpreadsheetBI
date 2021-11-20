let

    tbl = Table.FromRecords({
        [Primary Key = "blah", Full Description = "This is blah"], 
        [Primary Key = "hello", Full Description = "This is hello"], 
        [Primary Key = "zzzz", Full Description = "This is zzzz"]
        }), 

    ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"Primary Key", type text},
            {"Full Description", type text}

        })

in
    ChangedType