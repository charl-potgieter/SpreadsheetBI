#table(
 type table
    [
        #"Date"=date, 
        #"Description"=text,
        #"SubCategory"=text,
        #"Amount"=number
    ], 
 {
  {#date(2016,1,31), "blah","A", 1234},
  {#date(2017,2,28), "hello","A", 100},
  {#date(2016,1,31), "blah","A", 13334},
  {#date(2018,4,30), "hello","B", 1550},
  {#date(2016,1,31), "zzzz","A", 1034},
  {#date(2017,2,28), "hello","A", 1500},
  {#date(2016,1,31), "zzzz","A", 1734},
  {#date(2018,4,30), "hello","B", 10},
  {#date(2019,8,28), "blah","B", 1454},
  {#date(2020,5,31), "hello","B", 1560}
 }
)