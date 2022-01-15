# Stock Analysis

## I. Overview of Project

### Background
Steve has been researching about green stocks to help his parents in their investment decisions and the workbook we prepared for him proved very useful. But for his next step, he then wants to analyze the entire stock market over the last few years.  

### Objective
Since the workbook we provided for him was set-up to analyze just 12 stocks, we need to refactor the code in order to collect the same information much faster, so that it can be applicable for analyzing thousands of stocks. 


## II. [Analysis and Results](VBA_Challenge.xlsm)

### A. Results
![](Resources/VBA_Challenge_2017.png)

**All of the green stocks except for TERP had a positive return** in 2017 .

![](Resources/VBA_Challenge_2018.png)

But in 2018, **only 2 stocks (ENPH and RUN) had a positive return**. Looking for investments in other sectors is highly recommended

### B. Refactored Code
The 4 major steps taken in refactoring the code are shown below:
![](Resources/Refactored_Code.png)

The most significant changes in the code are:
- We declared the Starting Price and Ending Price **variables as Single (instead of Double)** which is more applicable since the values only have 2 decimals. The variables now hold less data.
- Instead of searching for one ticker data 12 times, **we searched the 12 tickers' data in one pass**.

### C. Execution Time
Comparison for 2017

![](Resources/Execution_Time_Comparison_2017.png)

Comparison for 2018

![](Resources/Execution_Time_Comparison_2018.png)

**The execution times of the two versions show that the refactored script is almost 5 times faster than the original script.**


## III. Summary
### A. Advantages / Disadvantages of Refactoring Code
**Advantages** 
- Refactoring improves the code by increasing efficiency and removing redundant moves leading to faster results. 
- The code will look cleaner once refactored, making it easier to understand .
- Refactoring is faster than rewriting the whole code. 

**Disadvantages**
- Refactoring can introduce bugs, and should be done by a skilled person. 


### B. The Original Vs Refactored VBA script
**Refactored Script**
- From our results, the refactored script runs faster than the original.
- The user can learn a lot from increasing efficiency of the code.

**Original Script** 
- The original script was so much easier to make and is a good starting point.
