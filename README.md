# Excel Formula

## Formula list
| T | Formula                                                               | Note                   |
|---|-----------------------------------------------------------------------|------------------------|
| 1 | ```=VLOOKUP(MAX(A1:A6),A1:B6,2,FALSE)```                              | Using Max with vlockup  |
| 2 | =INDEX($A$10:$A$15,MATCH(MIN(B10:B15),B10:B15,0))                     | Using Index with Match |
| 3 | =INDEX(A3:M484,MATCH(MAX(K3:K484),K3:K484,0),11)                      | Index with Match       |
| 4 | =INT(U2)+((MID(U2,3,3)/60)*100)+((U2-LEFT(U2,(LEN(U2)-5)))/3600)*10^4 | Computing cells        |
| 5 |


## Tips and Tricks

1. INDEX-MATCH: The INDEX-MATCH function is used to return a value in a column
   to the left. With VLOOKUP, you're stuck returning an appraisal from a column
   to the right. Another reason to use index-match instead of VLOOKUP is that
   VLOOKUP needs more processing power from Excel. This is because it needs to
   evaluate the entire table array which you've selected. With INDEX-MATCH,
   Excel only has to consider the lookup column and the return column [Ref. 2.].

2.

## References
1. [Index, Match and Vlookup Examples](http://www.contextures.com/xlFunctions03.html)
2. [Excel tutorial excel basics formula](https://www.simplilearn.com/tutorials/excel-tutorial/excel-formulas)

