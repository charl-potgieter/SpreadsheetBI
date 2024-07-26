let

    tbl = Table.FromRecords({
        [NullHeader1 = "",
        NullHeader2 = "",
        NullHeader3 = "",
        NullHeader4 = "",
        NullHeader5 = "",
        NullHeader6 = "",
        NullHeader7 = "",
        NullHeader8 = "",
        NullHeader9 = "",
        NullHeader10 = ""]
        }), 

    ChangedType = Table.TransformColumnTypes(
       tbl, 
        {
            {"NullHeader1", type text},
            {"NullHeader2", type text},
            {"NullHeader3", type text},
            {"NullHeader4", type text},
            {"NullHeader5", type text},
            {"NullHeader6", type text},
            {"NullHeader7", type text},
            {"NullHeader8", type text},
            {"NullHeader9", type text},
            {"NullHeader10", type text}
        })

in
    ChangedType