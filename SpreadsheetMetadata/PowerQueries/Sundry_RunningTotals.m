// Returns a list of running totals for the Values paramter
// Inspired by https://www.myonlinetraininghub.com/quickly-create-running-totals-in-power-query

(Values as list)=>
let
    BufferedValues = List.Buffer(Values),
    
    fn_Seed = () =>[Counter=0, RunningTotal=BufferedValues{0}],

    fn_ContinueWhileTrue = (CurrentRecord)=> CurrentRecord[Counter] <= (List.Count(BufferedValues) -1),

    fn_GenerateNextValue = (CurrentRecord)=>
    let
        NextRecord = [
            Counter = CurrentRecord[Counter] + 1,
            RunningTotal = CurrentRecord[RunningTotal] + BufferedValues{Counter}
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