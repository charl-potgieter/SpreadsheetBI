let
    Source = Table.Distinct(Table.SelectColumns(HeadingConsistencyData, {"Sub Folder Path"}))
in
    Source