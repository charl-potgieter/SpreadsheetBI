let
    Source = Table.Distinct(Table.SelectColumns(HeadingConsistencyData, {"Parent Path"}))
in
    Source