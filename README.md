# VBA-challenge

Database and instruction provided by UC Berkeley Extension Data Analytics Bootcamp.

Using VBA scripting to analyze real stock market data.

**Files**
- alphabetical_testing (test data) is used to develop my scripts.
- Mutiple_year_stock_data (stock data) is used to generate the final report for this assignment.
- A screenshot for each year of my results on the Multi Year Stock Data
- VBA Scripts as separate files

**Process & Credits**
- Scripts are written referencing then in class activities.
- I ran into the problem of not being identify the first open date, which was needed for getting the Yearly Change value.
  - I attended office hour, and my instructor Khaled showed me that I need to declare FirstOpen = 2 (it started on row 2) and add FirstOpen = i + 1 after so it will move on to the next year's first open date.
- I used this [website(https://excelhelphq.com/how-to-loop-through-worksheets-in-a-workbook-in-excel-vba/)]as reference for writing scripts that will loop through all of the worksheets.
- I ran into the problem of my scripts for conditional formatting the Yearly Change column incorrectly and my scripts that will convert Percent Change value to percentage didn't work either. I had to click run twice for the conditional formatting to work properly.
  - I asked for help through Data Bootcamp's Learning Assistant App on Slack, and the Learning Assistant suggested to me that the issue was because my condition if the value of the cell > 0 runs when there is no data in the cell. And he suggested me to create another for loop for the conditional formating and converting the values to percentage. I did as he suggested, and my problem was resolved.
- I used this [website] (https://excelvbatutor.com/vba_lesson9.htm) as reference to convert decimal value to percentage.
