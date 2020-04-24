let
    tbl = #table(
                type table
                    [
                        #"Date"=date, 
                        #"Description"=text,
                        #"SubCategory"=text,
                        #"Amount"=number
                    ], 
                {
                {#date(2016,1,31), "blah","A", 1234}
                }
                ),

    fn = 
    (state, current) =>
    let
        LastTableRow = Table.FromRecords({Table.Last(state)}),
        Output = Table.Combine({state, LastTableRow})

    in
        Output,

    acc= List.Accumulate({1..10}, tbl, fn)
in
    acc