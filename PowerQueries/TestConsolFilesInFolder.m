let

    //Function for debugging purposes
    fn = 
    (Folder, FileName) =>
    let
        Source = Excel.Workbook(File.Contents(Folder & FileName), null, true),
        tbl = Table.SelectRows(Source, each [Name] = "tbl_Data")[Data]{0},
        ChangedType = Table.TransformColumnTypes(tbl,{{"Field1", type text}, {"Field2", type text}})
    in
        ChangedType,

    Consol = fn_std_ConsolidatedFilesInFolder("C:\Users\charl\Dropbox\Dropbox_Charl\Computer_Technical\000_ExampleDataFiles\TestConsolFilesInFolder\YearlyWithDateInFile", fn, "YYYY", #date(2019,1,1), #date(2019,3,31), "MyDate")

in
    Consol