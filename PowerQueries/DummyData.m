#table(
 type table
    [
        #"Date"=number, 
        #"Description"=text,
        #"Amount"=number
    ], 
 {
  {#date(2016,1,31), "blah", 1234},
  {#date(2017,2,28), "hello", 100},
  {#date(2016,1,31), "blah", 1334},
  {#date(2018,4,30), "hello", 1050},
  {#date(2019,8,28), "blah", 144},
  {#date(2020,5,31), "hello", 150}
 }
)