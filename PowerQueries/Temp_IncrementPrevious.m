let
    tbl = Table.FromRecords({[
                            Date = #date(2016,1,31),
                            Description = "blah",
                            SubCategory = "A",
                            Amount = 1234
                            ]}),

    fn = (state, current) =>
    let
        PreviousTableRowAsRecord = Table.Last(state),

        NewRow = Table.FromRecords({[
                                    Date = PreviousTableRowAsRecord[Date], 
                                    Description = PreviousTableRowAsRecord[Description], 
                                    SubCategory = PreviousTableRowAsRecord[SubCategory], 
                                    Amount = PreviousTableRowAsRecord[Amount] + 1
                                    ]}),


        // AppendRow = Table.FromRecords({Table.Last(state)}),
        Output = Table.Combine({state, NewRow})

    in
        Output,

    acc= List.Accumulate({1..10}, tbl, fn)
in
    acc