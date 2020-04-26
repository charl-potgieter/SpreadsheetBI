/*-----------------------------------------------------------------------------------------------------------
    Returns a table of number scales for report formatting
-----------------------------------------------------------------------------------------------------------*/
#table(
    type table
    [
        #"ShowValuesAs" = text,
        #"DivideBy" = Int64.Type
    ],
    {
        {"CCY", 1},
        {"'000", 1000},
        {"m", 1000000}
    }
)