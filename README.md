# VBA-Challenge
 VBA scripting to analyze real, historical stock market data. Scripts will complete the following:
 
**_Part 1_**

Created for-loop to look through every line of data and create a summary table by calculating the following:
1. Create one line for each unique stock ticker value
2. Calculate the change in stock value from open on first day of year to close on last day of year.
    * Applied conditional formatting to highlight positive and negative change.
3. Figure the percentage change of stock value from open on first day of year to close on last day of year.
4. Calculate total stock volume over course of year.
5. Enclosed everything into a for-loop to calculate my code across all worksheets in the workbook.

**_Part 2_**

Used for-loops to cycle through newly created summary table to create additional summary table that highlighted the following:
1. Largest stock value increase and it's associated ticker.
2. Largest stock value decrease and it's associated ticker.
3. Largest stock volume and it's associated ticker.
4. Enclosed everything into a for-loop to calculate my code across all worksheets in the workbook.

**_Extra_**

While testing I found it inefficient to test my code and have to clear it when it would return an error. So, I picked three shapes to tie my macros to. Now, when you click one it will run the associated macro across all sheets in the workbook. Macros are mapped to the shapes as follows:
1. **White Wolf** - Runs the first part of my code that aligns with the first part of the homework assignment. This code creates the first summary table based on the raw stock data.
2. **Black Wolf** - Runs the second part of my code that aligns with the 'bonus' part of the assignment. This code creates the second, smaller summary table that is based on the table created in part one.
3. **Eraser** - This will run a macro that clears both formatting and values in both table ranges that are created by running either of the other macros.
