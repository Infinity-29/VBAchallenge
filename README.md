# VBA Homework - The VBA of Wall Street

## Background

You are well on your way to becoming a programmer and Excel master! In this homework assignment you will use VBA scripting to analyze real stock market data. Depending on your comfort level with VBA, you may choose to challenge yourself with a few of the challenge tasks.

### Before You Begin

1. Create a new repository for this project called `VBA-challenge`. **Do not add this homework to an existing repository**.

2. Inside the new repository that you just created, add any VBA files you use for this assignment. These will be the main scripts to run for each analysis.

### Files

* [Test Data](Resources/alphabetical_testing.xlsx) - Use this while developing your scripts.

* [Stock Data](Resources/Multiple_year_stock_data.xlsx) - Run your scripts on this data to generate the final homework report.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

## Instructions

* Create a script that will loop through all the stocks for one year and output the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* You should also have conditional formatting that will highlight positive change in green and negative change in red.

* The result should look as follows.

![moderate_solution](Images/moderate_solution.png)

### BONUS Challenge (Optional)

1. Your solution will also be	est % increase", "Greatest % decrease" and "Greatest total volume". The solution will look as follows:

![hard_solution](Images/hard_solution.png)

2. Make the appropriate adjustments to your VBA script that will allow it to run on every worksheet, i.e., every year, just by running the VBA script once.

### Other Considerations

* Use the sheet `alphabetical_testing.xlsx` while developing your code. This data set is smaller and will allow you to test faster. Your code should run on this file in less than 3-5 minutes.

* Make sure that the script acts the same on each sheet. The joy of VBA is to take the tediousness out of repetitive task and run over and over again with a click of the button.

* Make sure you save your workbook as a macro enabled workbook (.xlsm), or you VBA code may be lost.

* You can assume that entries for any given ticker are in consecutive rows and that the entries are in chronological order.

* Review the rubric (link below)

## Submission

* To submit please upload the following to Github:
  
  * A screen shot for each year of your results on the Multi Year Stock Data, similar to results images above.

  * VBA Script as a separate file. Use Export from the File Menu to save your script for submission.

  * DO NOT PUSH THE Excel Workbook to GitHub. (Hint: Add line *.xls to your .gitignore file)

* Ensure your repository has regular commits (i.e. 20+ commits), a thorough README.md file

* After everything has been saved, copy link for your GitHub repository for submission in BootCampSpot.

## Rubric

[Unit 2 Rubric - VBA Homework - The VBA of Wall Street](https://docs.google.com/document/d/1OjDM3nyioVQ6nJkqeYlUK7SxQ3WZQvvV3T9MHCbnoWk/edit?usp=sharing)

- - -

© 2021 Trilogy Education Services, LLC, a 2U, Inc. brand. Confidential and Proprietary. All Rights Reserved.
