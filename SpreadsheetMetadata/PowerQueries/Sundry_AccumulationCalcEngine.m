let 


//------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//      Function code
//------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

 
    FunctionExDocumenation =  
    (tbl_Input as table, TableRowCalcFunction as function)=>
    let


        fn_InsertRecordIntoRecordOfLists = 
        (ExistingRecordOfLists, RecordToInsert)=>
        let
            /*
                fn_InsertRecordIntoRecordOfLists inserts a record into a record of lists.The working of function best described by means of example:
                fn_InsertRecordIntoRecordOfLists(
                    [A = {1,2,3}, B = {5,6,7}, C = {9,10,11}],
                    [A = 4, B = 8, C = 12])
                Returns
                    [A = {1,2,3,4}, B = {5,6,7,8}, C = {9,10,11,12}]
            */

            fn_Accumulator = (state, current)=>   // Where state is the record of lists, current is the current field name
            let
                FieldName = current,
                CurrentList = try Record.Field(ExistingRecordOfLists, FieldName) otherwise {},
                ValueToAdd = Record.Field(RecordToInsert, FieldName),
                NewList = CurrentList & {ValueToAdd},
                RemovePreviousFieldInRecord = try Record.RemoveFields(state, FieldName) otherwise state,
                AddNewField = Record.AddField(RemovePreviousFieldInRecord, FieldName, NewList)
            in
                AddNewField,

            FieldNames = Record.FieldNames(RecordToInsert),    
            ReturnValue = List.Accumulate(FieldNames, ExistingRecordOfLists, fn_Accumulator)

        in
            ReturnValue,



        SeedRecord = [InputTable = Table.Buffer(tbl_Input)],
        IndexList = List.Buffer({0..Table.RowCount(SeedRecord[InputTable])-1}),

        fn_Accumulator = 
        (RecordOfLists, CurrentIndex)=> 
        let  
            CurrentInputRow = RecordOfLists[InputTable]{CurrentIndex},
            CurrentRecord = TableRowCalcFunction(tbl_Input, CurrentIndex, CurrentInputRow, RecordOfLists),        
            ReturnValue = fn_InsertRecordIntoRecordOfLists(RecordOfLists, CurrentRecord) 
        in
            ReturnValue,

        Accumulate = List.Accumulate(IndexList, SeedRecord, fn_Accumulator),
        RemoveInputTable = Record.RemoveFields(Accumulate, "InputTable"),
        ConvertToTable = Table.FromColumns(Record.ToList(RemoveInputTable), Record.FieldNames(RemoveInputTable))
    in
        ConvertToTable,



//------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
//      Add documentation metadata to the function
//------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------


    DocumentationMetaData = [
        Documentation.Name =  " Accumulation calculation engine ", 
        Documentation.LongDescription = "Calculation function to enable Excel style table calculations that reference previos data rows <br><br> " & 
            "tbl_Input contains all input data into the calculation <br><br>" & 
            "TableRowCalcFunction meets below criteria : <br>" & 
            " (1) returns a record <br>" & 
            " (2) has the below parameters: <br>" &  
            "   * InputTable as table, <br>" & 
            "   * CurrentIndex as number, <br>" & 
            "   * CurrentInputRow as record, <br>" & 
            "   * RecordOfLists as record) <br><br>" &  
            " For a more detailed explanation refer www.tba.... <br>" & 
            " Author: Charl Potgieter",
        Documentation.Source = "Source is TBA", 
        Documentation.Author = "Charl Potgieter"
    ],

    typeFunctionWithDocumentation = type function (
          tbl_Input as (type table),
          TableRowCalcFunction as function
            )
        as table meta DocumentationMetaData,

    // Replace the extisting type of the function with the individually defined
    Result =  Value.ReplaceType(FunctionExDocumenation, typeFunctionWithDocumentation)
 
 in 

    Result