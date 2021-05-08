# Stock-Analysis

## Overview of Project
Creating a Macro which enables our client to look at various stocks data throughout the year.
### Background
Our client Steve wants to find the total daily volume and yearly return for each stock. Daily volume is the total number of shares traded throughout the day; 
it measures how actively a stock is traded. The yearly return is the percentage difference in price from the beginning of the year to the end of the year.
### Purpose
Purpose of this project is to refactor an existing Macro that we created for stock analysis and return the results much quicker,
thus improving the performance of our code. Also using formatting we are showing positive returns in green color and negative returns in red.
## Results
### 2017
The Excel screenshot below shows the Stocks data and code performance for the year 2017.
* With original Macro the code was taking 0.80 seconds to run
* After refactoring the Macro code is taking 0.13 seconds to run
* Out of 12 stocks only TERP stock returned negative value and rest all performed well
* DQ gave highest return of 199.4%
___
![VBA_Challenge_2017](https://github.com/ysbcode/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png?raw=true)
___
### 2018
The Excel screenshot below shows the Stock data and code performance for the year 2018.
* With original Macro the code was taking 0.81 seconds to run
* After refactoring the Macro code is taking 0.12 seconds to run
* Out of 12 stocks only two stocks returned positive value ENPH gave highest return of 81.9% followed by RUN with 84.0%
* Last year's top performing stock DQ gave the least return with negative -62.6%
___
![VBA_Challenge_2018](https://github.com/ysbcode/stock-analysis/blob/main/Resources/VBA_Challenge_2018.png?raw=true)
___
## Summary
### Advantage and Disadvantages of Refactoring code
Refactoring is a key part of coding and most of the programmers need to refactor code and work on a code written by someone else. 
Sometimes when it is not commented or indented properly it could be frustrating and hard to link different pieces together. 
But sometimes, refactoring code will be an entry point to working with the existing code at a job. Main purpose of refactoring code is to make the code
more efficient, following fewer steps, using less memory or improving logic of the code to make it easier to understand for future users.
### Advantages and Disadvantages of the original and refactored VBA script
Original VBA script was good for analyzing a dozen stocks but it might not work as well for thousands of stocks. Even if it does, it may take a long time to execute.
Original VBA script followed simple logic and two For loops and with refactored script we included more For loops and even used a loop to display output values.
Refactored script is taking considerably lesser time to run the code and have improved code performance and efficiency.
