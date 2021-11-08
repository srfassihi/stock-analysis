# Stock Analysis

## Overview

### Purpose
- Automate Stock Analysis for a spreadsheet using VBA Macros
- Refactor existing code to run faster than the original script

## Results

### Original 'green_stocks' performance
- Runtime between 0.591 to 0.649 seconds

### Refactored code performance
- Runtime between 0.633 to 0.635 seconds
![2017](https://github.com/srfassihi/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

![2018](https://github.com/srfassihi/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)

### Outputs
![2017 Stock Analysis](https://github.com/srfassihi/stock-analysis/blob/main/Resources/2017%20Stocks.png)

![2018 Stock Analysis](https://github.com/srfassihi/stock-analysis/blob/main/Resources/2018%20Stocks.png)

## Conclusions
### Advantages of Refactoring
- Refactoring existing code and adding comments can make it more readable and performant.
- Following consistent design patterns can help others understand the code and make it more maintainable.
- Reducing hard coding can potentially make the code more reusable.

### Disadvantages of Refactoring
- Risk of breaking existing code and creating regressions.
- Significant additional time spent on refactoring code can increase a project's deadline.

## Original vs. Refactored Script
- Not a significant difference in overall performance from original script to new script due to:
	- Similar looping through rows
	- Similar variables
	- Formatting script included in the Refactored script but not in the original
