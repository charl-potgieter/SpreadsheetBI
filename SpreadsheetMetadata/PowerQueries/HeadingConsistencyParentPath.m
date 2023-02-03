let
    Source = Table.Distinct(Table.SelectColumns(HeadingConsistencyData, {"Parent Path"}), Comparer.OrdinalIgnoreCase)
in
    Source