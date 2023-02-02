let
    Source = Table.Distinct(Table.SelectColumns(HeadingConsistencyData, {"Field Name"}))
in
    Source