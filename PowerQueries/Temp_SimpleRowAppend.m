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

    acc= List.Accumulate({1..10}, tbl, (state, current)=> Table.Combine({state, tbl}))
in
    acc