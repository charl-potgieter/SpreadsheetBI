/*
    Returns true if data access query should be loaded:
    (1) Checks which reports require refresh
    (2) Gets the list of data access queries for those reports
    (3) Returns true if QueryName is in above list
*/

(QueryName as text)=>
let
    DataQueriesPerReport = Excel.CurrentWorkbook(){[Name="tbl_DataAccessQueriesPerReport"]}[Content],
    AllReports = Excel.CurrentWorkbook(){[Name = "tbl_ReportList"]}[Content],
    ReportsForRefreshTable = Table.SelectRows(AllReports, each [Run with table refresh] <> null),
    ReportsForRefreshList = ReportsForRefreshTable[Report Name],
    FilterQueriesBasedOnReport = Table.SelectRows(DataQueriesPerReport, each List.Contains(ReportsForRefreshList, [Report Name])),
    ReturnValue = List.Contains(FilterQueriesBasedOnReport[Data Access Query Name], QueryName)
in
    ReturnValue