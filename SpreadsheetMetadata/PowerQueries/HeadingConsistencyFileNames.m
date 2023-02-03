let
    Source = Table.Distinct(Table.SelectColumns(HeadingConsistencyData, {"File Name"}), Comparer.OrdinalIgnoreCase)
in
    Source