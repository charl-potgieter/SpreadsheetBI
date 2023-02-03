let
    Source = Table.Distinct(Table.SelectColumns(HeadingConsistencyData, {"Sub Folder Path"}), Comparer.OrdinalIgnoreCase)
in
    Source