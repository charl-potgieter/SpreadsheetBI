let
    Source = Table.Distinct(Table.SelectColumns(HeadingConsistencyData, {"Field Name"}), Comparer.OrdinalIgnoreCase)
in
    Source