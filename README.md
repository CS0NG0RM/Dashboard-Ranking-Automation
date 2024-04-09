# Dashboard-Ranking-Automation

## Overview

This project consists of several VBA modules that automate the process of updating the ranking of occupations in an Excel workbook. Each module retrieves the occupation and ranking from specific sheets, finds the corresponding row in the data sheet where the occupation matches, and inserts the ranking into the appropriate row and column.

The end goal is to ensure that a rating can be inputted in each sheet that has a Dashboard. The data should then filter back into the original dataset where the rankings can be stored and retrieved at any point if needed.

## Highlights
### Efficient Data Manipulation
The modules demonstrate efficient data manipulation by retrieving and updating data across multiple sheets. They showcase the use of Excel's `WorksheetFunction.Match` to find the corresponding row for an occupation in the data sheet.

## Modules

### Module 1
The first module retrieves the occupation and ranking from the "Dashboard" and "Occupation" sheets, respectively. It then finds the corresponding row in the "Data" sheet where the occupation matches and inserts the ranking into the appropriate row in column "CR".

### Dynamic Data Insertion
The modules dynamically insert the ranking into different columns ("CR" to "CX") in the data sheet, demonstrating flexibility in handling data structures.

### Modular Design
The design of the project is modular, with each module performing a specific task. This makes the code easier to understand, test, and maintain.

## Running the Code
The `SaveRanking` subroutine calls the `UpdateRanking` subroutine in each module sequentially, demonstrating how to orchestrate multiple tasks.

