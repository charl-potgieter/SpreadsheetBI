let
    Source = Table.Distinct(Table.SelectColumns(HeadingConsistencyData, {"File Name"}))
in
    Source