# VBA-challenge

Database and instruction provided by UC Berkeley Extension Data Analytics Bootcamp.

Using VBA scripting to analyze real stock market data.

**Files**
- alphabetical_testing (test data) is used to develop my scripts.
- Mutiple_year_stock_data (stock data) is used to generate the final report for this assignment.
- A screenshot for each year of my results on the Multi Year Stock Data
- VBA Scripts as separate files

This was my first time writing scripts in VBA (also my first time doing coding-like task), and I struggled to complete this assignment. I did a lot of website searching/googling, and asking for help in office hour and through Learning Assistant App. I will list the website I used as reference and the help I received in "Process & Credits."

**Process & Credits**
- Scripts are written referencing then in class activities.
- I ran into the problem of not being identify the first open date, which was needed for getting the Yearly Change value.
  - I attended office hour, and my instructor Khaled showed me that I need to declare FirstOpen = 2 (it started on row 2) and add FirstOpen = i + 1 after so it will move on to the next year's first open date.
- I used this website(https://excelhelphq.com/how-to-loop-through-worksheets-in-a-workbook-in-excel-vba/) as reference for writing scripts that will loop through all of the worksheets.
- I ran into the problem of my scripts for conditional formatting the Yearly Change column incorrectly and my scripts that will convert Percent Change value to percentage didn't work either. I had to click run twice for the conditional formatting to work properly.
  - I asked for help through Data Bootcamp's Learning Assistant App on Slack, and the Learning Assistant suggested to me that the issue was because my condition if the value of the cell > 0 runs when there is no data in the cell. And he suggested me to create another for loop for the conditional formating and converting the values to percentage. I did as he suggested, and my problem was resolved.
- I used this website (https://excelvbatutor.com/vba_lesson9.htm) as reference to convert decimal value to percentage.
- I ran into the problem of my conditional formatting and converting value to percentage not working on only the last sheet. When I debugged, it highlighted the PercentChange = YearlyChange/FirstOpen.
  - I asked for help though the Learning Assistant App again, the Learning Assistant suggested to me that it was a divide by zero error. Before a workable solution was suggested by the Learning Assistant, I was able to look up the issue and found the solution on this website (https://www.mrexcel.com/board/threads/help-with-avoiding-division-by-zero-error-in-vba.783862/). I placed an if statement that states if FirstOpen <> 0 then.
- I ran into another problem when I was creating scripts for the Greatest % Change. I used this website (https://www.excelanytime.com/excel/index.php?option=com_content&view=article&id=105:find-smallest-and-largest-value-in-range-with-vba-excel&catid=79&Itemid=475) as reference for using worksheet function max. But I didn't know how to define the last row by using a find last row formula that I already created (lastroww). I didn't want to hard find the last row, and just in case all the worksheets' last row numbers are different. So I tried using for look to find the greatest value in Percent Change column, but it ended up only giving me the last value on that column.
  - I asked the Learning Assistant App for help again after struggling for a while. The Learning Assistant suggested I use the worksheet function for finding the max value and also provided me a function that will use my lastroww formula. The solution he gave me is GreatestPercent = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & lastroww)), ws.Range("K2:K" & lastroww), 0), and since it gave me the row number where that value is located starting from the second row (where the first value is), so I had to + 1 to the result in a separate cells declaration. To better understand the worksheet function match, I used this website (https://www.wallstreetmojo.com/vba-match/) as reference.
