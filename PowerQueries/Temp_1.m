let
    Source = List.Accumulate({1..500}, 0, (state, current)=> state+current)
in
    Source