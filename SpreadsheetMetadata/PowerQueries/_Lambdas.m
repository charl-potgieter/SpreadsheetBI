let

    fn_GetFileContents = 
    (GitUrl)=>
    let
        Source = Text.FromBinary(Web.Contents(GitUrl)),
        SplitText = Text.Split(Source, "//////////"),
        ConvertedToTable = Table.FromList(SplitText, Splitter.SplitByNothing(), {"Combined"}, null, ExtraValues.Error),
        PositionOfColon = Table.AddColumn(ConvertedToTable, "PositionOfColon", each Text.PositionOf([Combined], ":"), Int64.Type),
        AddItemCol = Table.AddColumn(PositionOfColon, "Item", each Text.Start([Combined], [PositionOfColon]), type text),
        AddValueCol = Table.AddColumn(AddItemCol, "Value", each Text.End([Combined], Text.Length([Combined]) - [PositionOfColon] -2), type text),
        SelectCols = Table.SelectColumns(AddValueCol,{"Item", "Value"}),
        CleanItemCol = Table.TransformColumns(SelectCols,{{"Item", Text.Clean, type text}}),
        TransposeTable = Table.Transpose(CleanItemCol),
        PromotedHeaders = Table.PromoteHeaders(TransposeTable, [PromoteAllScalars=true]),
        ChangedType = Table.TransformColumnTypes(PromotedHeaders,{{"Author", type text}, {"Description", type text}, {"PipeDelimitedParameterAndDescription", type text}, {"RefersTo", type text}})
    in
        ChangedType,


    Source = Json.Document(Web.Contents("https://api.github.com/repos/charl-potgieter/ExcelLambdas/contents")),
    ConvertToTable = Table.FromList(Source, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    RenamedColumn = Table.RenameColumns(ConvertToTable,{{"Column1", "Records"}}),
    AddTypeCol = Table.AddColumn(RenamedColumn,"Type", each [Records][type], type text),
    FilterOnDirectories = Table.SelectRows(AddTypeCol, each ([Type] = "dir")),
    AddCategoryColumn = Table.AddColumn(FilterOnDirectories, "Category", each [Records][name], type text),
    AddFileUrlColumn = Table.AddColumn(AddCategoryColumn, "url", each [Records][url], type text),
    AddListOfUnderlyingFiles = Table.AddColumn(AddFileUrlColumn, "RecordsOfFormulaFiles", each Json.Document(Web.Contents([url]))),
    Expanded = Table.ExpandListColumn(AddListOfUnderlyingFiles, "RecordsOfFormulaFiles"),
    AddFormulaNameCol = Table.AddColumn(Expanded, "Name", each Text.Start([RecordsOfFormulaFiles][name], Text.Length([RecordsOfFormulaFiles][name])-4), type text),
    AddDownloadUrlCol = Table.AddColumn(AddFormulaNameCol, "download_url", each [RecordsOfFormulaFiles][download_url]),
    AddGitUrlCol = Table.AddColumn(AddDownloadUrlCol, "GitUrl", each [RecordsOfFormulaFiles][_links][html], type text),
    AddFileContentsTable = Table.AddColumn(AddGitUrlCol, "FileContents", each fn_GetFileContents([download_url]), type table),
    ExpandedFileContents = Table.ExpandTableColumn(AddFileContentsTable, "FileContents", {"Author", "Description", "PipeDelimitedParameterAndDescription", "RefersTo"}, {"Author", "Description", "PipeDelimitedParameterAndDescription", "RefersTo"}),
    SelectCols = Table.SelectColumns(ExpandedFileContents,{"Category", "Name", "Author", "Description", "PipeDelimitedParameterAndDescription", "RefersTo", "GitUrl"})
in
    SelectCols