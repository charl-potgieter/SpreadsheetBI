/*
    Generates a dummy table.
    Utilised as a table for storage of  measures that can be utilised to form a basis for DAX table queries
    that can be output directly to Excel
*/
   

#table(
        type table[
            #"DummyFieldName"=text 
            ], 
        {
            {null}
        }
    )