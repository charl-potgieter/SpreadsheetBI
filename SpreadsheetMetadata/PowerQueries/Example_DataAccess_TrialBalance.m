(LoadData as logical)=>
let
    FolderPath = "C:\Users\charl\Dropbox\Dropbox_Charl\Computer_Technical\Programming_GitHub\SpreadsheetBI\ExampleDataFiles\TrialBalance",
    tblRaw = fn_std_ConsolidatedFilesInFolder(FolderPath, fn_std_Single_PipeDelimitedText, LoadData, null, null, null)
in
    tblRaw