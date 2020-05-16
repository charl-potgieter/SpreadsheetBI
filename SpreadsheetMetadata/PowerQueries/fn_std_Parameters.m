let


    // -----------------------------------------------------------------------------------------------------------------------------------------
    //                      Documentation
    // -----------------------------------------------------------------------------------------------------------------------------------------
    
    Documentation_ = [
        Documentation.Name =  " fn_std_Parameters", 
        Documentation.Description = " Returns parameter value set out in  tbl_Parameters" , 
        Documentation.LongDescription = "  Returns parameter value set out in  tbl_Parameters", 
        Documentation.Category = "Text",  
        Documentation.Author = " Charl Potgieter"
        ],


    // -----------------------------------------------------------------------------------------------------------------------------------------
    //                      Function code
    // -----------------------------------------------------------------------------------------------------------------------------------------

    fn_=
    (parameter as text)=>
    let
        Source = Excel.CurrentWorkbook(){[Name = "tbl_Parameters"]}[Content],
        FilteredRows = Table.SelectRows(Source, each [Parameter] = parameter),
        ReturnValue = FilteredRows[Value]{0}
    in
        ReturnValue,




// -----------------------------------------------------------------------------------------------------------------------------------------
//                      Output
// -----------------------------------------------------------------------------------------------------------------------------------------

    type_ = type function (
        parameter as (type text)
        )
        as text meta Documentation_,

    // Replace the extisting type of the function with the individually defined
    Result =  Value.ReplaceType(fn_, type_)
 
 in 
    Result