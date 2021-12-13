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

# Day 8
Decoding seven segment displays. Part 1 was fairly easily accomplished with SUMPRODUCT() and a double unary. Part 2 is where the challenges started. Excel does not have a native character sorting function. I created one using LAMBDA() to eliminate mountainous formulas just to arrange the letters within a cell. The next challenge came from Excel's inability to use nested LAMBDA() functions. This required the use of multiple helper columns to sort some output for comparison with XLOOKUP(). Seven segment displays have a certain set of rules based on knowing the wiring for 4 of the 10 combinations. This breaks down into two blocks, one with 6 characters and one with 5 characters. Within 6 characters the pattern that matches 4 inside must be 9, the pattern the matches 1 inside must be 0, and the remainder is 6. This is represented in column AH to AJ with a LAMBDA() array loop. Within the 5 character sets the pattern that matches 1 inside must be 3, the one that is off by 1 from 6 is 5, and the remainder is 2. The checks for character matching do not require a sort but require data to be pared down from the full 10 group hence the two 3 cell groups for 6 character lengths and 5 character lengths. After generating the "decryption key" the output column is sorted and then XLOOKUP() across, CONCAT(), and then INT() to correct the value type.

# Day 9
Mapping sea vents. Part 1 is relatively straight forward. I took a slightly longer route of making bespoke formulas for each corner and the sides of the map for part 1. Part 2 required a bit more trickery and a run off box allowed for a single LAMBDA() to do most of the heavy lifting. I started by giving a unique number to each cell then iterated through the entire 100x100 grid square taking the lowest number in a plus shape around a target cell. Repeat this a few times until the sum of two 100x100 grids is the same then some simple array math cranks out a couple tables. I think I over did the last bit by getting the original tables back to show the final areas found but this could be shortcutted with a different lookup.

# Day 10
Invalid syntax. Part 1 requires running a generational LAMBDA() across the input until there the erroneous character is found. After singling out this character it is XLOOKUP() into a score. Thank goodness they didn't have us go back and find the position of this in the original string. The score is then summed at the bottom. Part 2 takes the remained of the lines using a LOOKUP() to remove blanks and then sort them to push blanks to the bottom. This output then has a character by character replacement to the closing character which is then reversed and exploded across an array. These columns are then XLOOKUP() to turn into scores and then two formulas do the math to solve across. A check formula is added to ensure that there is no funny business and the MEDIAN() matches an offset.

# Day 11
Glowing octopus. Part 1 required building a careful formula to look at the previous values and generationally map out when an octopus' energy level exceeded 10 and then zero it out for the remainder of that round. Because laziness is king each grid is surrounded by Xs as runoff so the formula within each cell in the grid is identical.  Due to the variable length of generations there is a bit of extraneous computations for shorter generations but long enough to capture all the corner cases. At the end of each row the number of zeros in the box is calculated for later sum at Step 100. Pt 2 is simply paste it on down some more until you get a grid of all 0s. Depending on the input this can occur at different step. A SUM()=0 check identifies the winning grid and the step number it occured on.
Edit: I discovered there is a MAP() function now. Instead of being 23MB it's now closer to ~4MB and should work with any input.

# Day 12
Something something path finding with caves and nonsense. This is where Excel kind of falls apart. I haven't done this one yet.

# Day 13
Folding paper. Theory has it you can't fold a single sheet of paper more than 8 times before it spontaneously combusts but in this world you can fold it 12 times. Part 1 consisted of finding the formula of 2y-x and applying to the input. This yields a field of dots that needs to be checked for unique dots and then counted. Part 2 requires completing the 12 folds and then generating the field of dots to reveal the hidden message. A nice touch indeed.
