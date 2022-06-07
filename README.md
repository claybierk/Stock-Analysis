# Stock Analysis Using VBA

## Overview

The objective of this analysis was to refactor VBA code so that it could be optimized to produce the desired output with more efficiency. This was accomplished by writing code to loop through all the data at once versus one cell at a time.

#### Skills Used:
- FOR loops
- IF-THEN statements
- Refactoring VBA code

## Results

2017 Refactored:

<img width="577" alt="Output+Time17" src="https://user-images.githubusercontent.com/100488626/172439096-e166ced7-012a-4b85-a963-b579cfe32235.png">

2017 Before: 

<img width="262" alt="2017_Base" src="https://user-images.githubusercontent.com/100488626/172440683-b8907e0e-fbdf-44f2-ae4d-f4f4e83c6cd1.png">

2018 Refactored:

<img width="577" alt="Output+Time18" src="https://user-images.githubusercontent.com/100488626/172439131-549a20b3-ae0e-46cf-9f68-8c16b14a87d1.png">

2018 Before:

<img width="262" alt="2018_Base" src="https://user-images.githubusercontent.com/100488626/172440723-1f2d6c43-5248-4078-8965-640d6c5b8d71.png">

#### Stock Performance Comparison

2017 was a much better year for these stocks with only 2 managing to stay in the green both years (ENPH & RUN). It appears 2018 was a tough year on all stocks, but RUN boosted its return suggesting a successful product launch or lucrative contract to buck the downward trend of the market. Based on some of the returns for 2018, I think there had to have been a major market correction or catastrophe to cause such carnage after a good year in 2017 in which 4 of our selected stocks doubled their market values (DQ, EPNH, FSLR, SEDG).

#### Execution Times Comparison

Refactoring made a huge difference in the overall time it took to execute our analysis, resulting in roughly 1/5 of the time being shaved off.

## Summary

#### Advantages

Refactoring code has 3 main advantages:
  1. Using less lines of code to accomplish the same result
  2. Improving the flow to make it more logical to follow
  3. Using less memory

#### Disadvantages

Refactoring can be very time consuming and thus lead to a lot of frustration or confusion especially if the original code isn't your own. Additionally, since it is done to working code to improve the things listed above, you could mess up the source code and possibly disrupt the functionality.

#### Decision

In this case there is no question that refactoring was the correct decision. Given the relative simplicity of this analysis, it made a lot of sense to refactor and improve the flow and functionality of this code.
