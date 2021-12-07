# aoc-excel-2021

Excel is programming language. This repo archives my attempt to do all of the Advent of Code 2021 entirely in Excel without the use of VBA.

# Day 1
walk in the park. some simple IF statements and SUMs generate the answer.

# Day 2
input is formatted with text to columns and then processed with some IF statements doing addition and multiplication and then summed at the bottom to generate the requsite answers.

# Day 3
things start to get tricky here. Excel has issues with zero padded numbers requiring an additional step to replace zeros lost on input. Next step is to create seperated helper columns to simplify part 1. They could be eliminated in favor of some additional MID() in the lower COUNTIF()s. Part 2 requires using the information from part 1 to walk through the input and remove items based on bits in the middle. Array formulas used here reduce calculation time as well as improving readability. The second issue excel faces here is the limit of 10 bits for BIN2DEC(). By left padding to 16 bits and then splitting apart allows for calculation to decimal and then the multiplication needed for the answers

# Day 4
it's time for bingo. Input is processed with text to columns and the rest is accomplished through formulas. Conditional formatting is added for brevity and to assist in finding the answer for part 2. Part 1 is automatically found while part 2 needs manual searching. I may revist this later and correct this. Boards are checked with XLOOKUP against a sliding window from the input stored in row 2. A MATCH() runs in row 605 to search for the first match in a row. It only supports Up, Down, Left, Right matches but not diagonals. Once found some offset trickery finds the top left corner and then returns the original board and the marked board. They are then compared and the sum is provided. Part 2 is similar but requires a manual copy paste of the entry to find the answer.

# Day 5
the big boy. raw computing power is used to crunch this beast of a problem. Input is split into helper columns and then a list of marked points are generated in the XpY format and spanned to the right. Part 1 filters only x and y moves that then feed into the monster 1000x1000 table. Part 2 performs some additional checks for / vs \ diagonals and generates the appropriate coordinates. Two 1000x1000 tables are built using COUNTIF() across the entire array and then a COUNTIF() across that table to generate the solutions. This sheet takes an extremely long time to compute due to the poor use of full text searches. Some form of intelligent hashing magic would speed this up but as I'm not gunning for the global leaderboard that task can wait.

# Day 6
more brute forcing. Part 1 was calculated without SUMPRODUCT() magic. Part 2 utilizes this starting from day 18 since for some reason it breaks below this bound. Additional processing of the input could reduce the vertical height of the table due to the limited number of unique patterns for the fish spawns.

# Day 7
Array formulas make this task a breeze. SEQUENCE() is used to save time but is not essential. Part 2 requires the development of a different formula for charting out all the possible combinations. In this instance the movement of our crab submarines can be graphed using 1/2n(2n+1). Once tables are computed fuel usage is summed at the bottom and and then MIN() is used to find the lowest value. A brute force approach that is elegant in it's simplicity.
