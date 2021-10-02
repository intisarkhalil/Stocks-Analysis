# Stocks Analysis:
## Overview of the Project:
This project gives analysis for green energy production, there are many forms of green energy to invest in, including: hydroelectricity, wind energy, geothermal energy, and bioenergy. The client decided to invest into **DAQO** New Energy Corporation, a company that makes silicon wafers for solar panels. DAQO’s ticker symbol is DQ. The researcher promised to investigate **DAQO** stocks, but he concerned about diversifying the client funds. The clint want to analyze a handful of green energy stocks in addition to DAQO’s stock. He created an **EXCEL** file containing the stock data. The initial analysis provides a well structure code to return the specific results. 
### Purpose:
Refactoring is the art of reworking your code into a more simplified or efficient form in a disciplined way. Refactoring improves internal code structure without altering its external functionality by transforming functions and rethinking algorithms. 
The main purpose of this project is to give the researcher the information about stock data in varies resources. The specific purpose of this project is to make the code more efficient, and to prepare a workbook, that, at the click of a button, the researcher can analyze an entire dataset. 
## Analysis and Challenges:
In this analysis, the researcher edits, or refactor, the solution code to loop through all the data one time to collect the same information, which is ensure that the refactoring of the code successfully made the **VBA** script run faster.  The process of the analysis to prepare a refactored code to run the yearly stock analysis including the following steps:
  - Create a subroutine named All Stocks Analysis.  
  - Define Start Time and End Time as Single variables
  - Define the year variable to input the year in the box message, using ``` InputBox() ``` function.
  - Assign the stater time to the ``` Timer ``` function.
  - Use ``` Range() ``` function to assign the title in cell A1.
  - Create a header row values using, ``` Cells() ``` function.
  - Initialize a ticker array as string, using ``` DIM ``` keyword
  - Get the number of row to loop over, using the following cod ``` RowCount = Cells(Rows.count, “A”).End(xlUP).Row ```
  - Create a ticker index variable as Single and initialize it by zero.
  - Create three output arrays using ``` Dim ``` keyword to assign the ticker volume as Long, and ticker starting and ending prices as Single
  -	Create a for loop to initialize the ticker volume to zero.
  -	Loop over all the rows in the spreadsheet.
    1. Increase the volume for current ticker.
    2. Check if the current row is the first row with the selected ticker Index.
    3. Check if the current row in the last row with the selected ticker Index, if the current the next row’s ticker doesn’t match, increase the ticker Index.
    4. Increase the ticker Index.
  -	Loop through array to output the Ticker, Ticker Volume and Return.
  -	Loop over formatting setting:
    - Assign font style to bold style through the range (A3:C3).
    - Add ``` borderbottom ``` line style, through the range (A3:C3) using ``` lcontinuous ```.
    - Change the number format for the range (B4:B15), using ``` (#,##0) ``` format.
    - Change the number format for the range (C4:C15), using ``` (0.0%) ``` format.
  - Loop over all value in column 3 to change the cell color according to condition that, if the cell value greater than zero (positive) then assign the cell color to green and  if not (negative) assign the color to red.
  - Finally, create the end time timer, and create a box message ``` MsgBox() ``` to present the elapsed run time for the refactored code. And end the Subroutine.
## Results:
Results of the **VBA** Refactored code can be pointed as follows:
  1. The refactor code return the yearly stocks analysis result for the year 2018, 2017, shown in the following image.
  2.	Results are well structure and readable, using **comments** and **whitespace**. 
  3.	Using font style, border bottom line style, number, color, and conditional formatting, lead to very interactive results layout.
  4.	looping through the data one time and collect all the information result in running the code so faster.
  5.	Also, the ``` Button ``` that use to run the codes are very useful. 
## Summary: 
Refactoring is the art of reworking your code into a more simplified or efficient form in a disciplined way. Refactoring improves internal code structure without altering its external functionality by transforming functions and rethinking algorithms. 
Using refactoring processing described above, on our codes we find the following conclusions: 
  1.	The code is well structure and easy to read.  
  The following figures she that the code is well structure and easy to read.
  
        ![A](https://user-images.githubusercontent.com/62036983/135702292-008706ba-b42b-408f-bd73-21c257db6e87.png)

        ![2](https://user-images.githubusercontent.com/62036983/135702323-6cb90663-ba2d-4494-b88d-85d1fedbd503.png)
        
        ![3](https://user-images.githubusercontent.com/62036983/135702334-9c8d64ed-bea8-493d-8f40-bfbf85241076.png)
        
  2. The running time is faster after factoring.

  The following image show the running time before the factoring process. 
  Figure a.1 (befor refactoring) for the year 2018 as an example. 
        ![a1](https://user-images.githubusercontent.com/62036983/135702385-a40bf78a-238b-418c-a21d-401d484ee944.png)

  Figure a.2 (after refactoring) for the year 2018 as an example
        ![Screenshot (127)](https://user-images.githubusercontent.com/62036983/135702408-53eb5216-147d-444c-b641-7606eb0ef461.png)
