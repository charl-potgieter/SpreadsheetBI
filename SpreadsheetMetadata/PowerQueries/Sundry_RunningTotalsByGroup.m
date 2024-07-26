// Returns a list of running totals for the Values paramter, resetting each time there is a change in GroupBy
// Inspired by https://www.myonlinetraininghub.com/grouped-running-totals-in-power-query

(Values as list,Grouping as list)=>
let
    BufferedValues = List.Buffer(Values),
    BufferedGrouping = List.Buffer(Grouping),
    
    fn_Seed = () =>[Counter=0, RunningTotal=BufferedValues{0}],

    fn_ContinueWhileTrue = (CurrentRecord)=> CurrentRecord[Counter] <= (List.Count(BufferedValues) -1),

    fn_GenerateNextValue = (CurrentRecord)=>
    let
        NextRecord = [
            Counter = CurrentRecord[Counter] + 1,
            RunningTotal = if BufferedGrouping{Counter} = BufferedGrouping{Counter - 1} then 
                    CurrentRecord[RunningTotal] + BufferedValues{Counter}
                else
                    BufferedValues{Counter}
        ]
    in
        NextRecord,

    fn_ReturnValue = (CurrentRecord)=>CurrentRecord[RunningTotal],

    Output = List.Generate(
        fn_Seed,
        fn_ContinueWhileTrue,
        fn_GenerateNextValue,
        fn_ReturnValue
    )
    
in
    Output