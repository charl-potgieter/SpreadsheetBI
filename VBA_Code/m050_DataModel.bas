Attribute VB_Name = "m050_DataModel"
Option Explicit



Sub CreateDaxQueryTable()

    Dim lo As ListObject
    
    Set lo = ActiveSheet.ListObjects.Add( _
        SourceType:=xlSrcModel, _
        Source:=ActiveWorkbook.Connections("Query - CreatedTable"), _
        Destination:=Range("$B$5"))
    
    lo.TableObject.WorkbookConnection.OLEDBConnection.CommandType = xlCmdDAX
    lo.TableObject.WorkbookConnection.OLEDBConnection.CommandText = "EVALUATE VALUES(CreatedTable[Text Column])"
    
    lo.TableObject.Refresh
    

End Sub



