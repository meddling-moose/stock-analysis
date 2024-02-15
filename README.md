# stock-analysis
My approach for this challenge was to separate each of the steps into manageable code snippets.

I started with a first pass at the pseudocode for the looping through the data on one sheet.
Starting with variable initialization, I declared my variables and for some, I set them to be specific values.
I created a loop through all the rows with data in the sheet and then I set values for the ticker, previous ticker, and next ticker.
Using the ticker information, I created three conditionals for whether we were currently on the beginning, middle, or end of a ticker run.
If the previous ticker is not equal to the current ticker:
- I set the volume and the opening price
- I increment the ticker count here because this should only happen when I find a new symbol
- The ticker count is a variable used to help place the ticker information in the aggregate columns such that when I'm populating the values, it doesn't just put it in the same row as the last row of a ticker symbol
If the ticker is equal to both the next ticker and the previous one:
- I add onto the count of volume
And if the ticker is not the same as the next ticker, then we have found the end of a run and I do a bunch of calculations and output some data:
- I record the closing price and finalize the volume calculation.
- I set the ticker value in the ticker column
- I calculate the yearly change and then I output that value in the appropriate column
- I calculate the percent change and then I output that in the appropriate column
- I ouput the total volume in the appropriate column
Finally, I end the row.

At this point, I feel pretty confident with my work, but I still need to add some increased functionality, so I go back through the script I wrote and I added more pseudocode for the areas that needed it:

The main features that need implementation are: 
- the column and row titles that need to be placed around the sheet
- the conditional formatting of the color behind the yearly change column
- the conditional formatting of the percent change column to be percentages with two decimal points of specificity
- the greatest percentage increase, greatest percentage decrease, and greatest total volume values
- looping through all of the sheets in the workbook automatically

I start with the column and row title ouptus. Thos were simply hardcoding values for where they were meant to go in the sheet.
Then I cover conditional formatting of color with a simple if statement that compares the yearly change value with zero and colors the cell accordingly.
Then I do the conditional formatting for the percentages with a simple format method.
For the greatest % increase, decrease, and total volume, I had to create some helper variables that store the current highest/lowest amounts and the corresponding tickers. During the 'end of a run' conditional, I compare the current ticker to these saved values, initialized ahead of the loop, and if they are more or less, then I replace the saved value with the new, greater value.
After the row ends, when it has looped through all of the data, I output the data into the respective table in excel.
The looping through all sheets was a little tricky because I had to go back through my macro and for every instance I called Cells(), I needed to add ws. to the front of it to let the macro know its for the aprticular worksheet. After that, it was just about writing a new loop to go outside the loop I'd written for a single sheet.

After this, I cleaned up some random bugs and made sure the variables that span multiple rows get cleared in time for the next sheet to start, and then I called it a day.

The data I got from this project was a combination of two excel spreadsheets, one with a ton of stock market data over three years, and the other is a smaller version of this filtered into multiple sheets to make testing easier. I got this data from my instructors and I don't believe it is real data, but could be wrong.
