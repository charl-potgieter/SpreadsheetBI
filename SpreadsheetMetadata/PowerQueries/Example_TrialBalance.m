let
    LoadData = true,
    TblRaw = Example_DataAccess_TrialBalance(LoadData),
    RemovedCols = Table.RemoveColumns(TblRaw,{"Name", "Folder Path"}),
    ChangedType = Table.TransformColumnTypes(RemovedCols,{{"EndOfMonth", type date}, {"Account Code", type text}, {"Amount", type number}})
in
    ChangedType