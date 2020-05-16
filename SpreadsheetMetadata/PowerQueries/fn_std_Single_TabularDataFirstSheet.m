(FolderPath as text, FileName as text)=>
let
    tbl = Excel.Workbook(File.Contents(FolderPath & FileName), true, null)[Data]{0}
in
    tbl