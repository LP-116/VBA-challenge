# VBA Challenge
## VBA Homework - The VBA of Wall Street

### Task
To use VBA scripting to analyse stock market data.
The VBA script needs to output the below information for each year of data _(each year is on a separate tab)_:
  *	The ticker symbol
  *	The difference between the opening price and the closing price for each ticker
  *	The percentage change for the price difference
  *	The total stock volume for each ticker

Conditional formatting also needs to be applied that highlights positive changes in green and negative changes in red.

For the bonus section of the task, the script needs to output the below information:
  *	The ticker with the greatest percentage increase
  * The ticker with the greatest percentage decrease
  *	The ticker with the greatest total volume

### Method overview

To generate a list of ticker symbols the code goes through the years data and identifies where a change occurs in the ticker name column. When a change occurs (i.e. the start of a new ticker is identified), the result is printed in the summary table section. The code compares the first ticker line to the line underneath it to identify a change in the ticker. Once the change is detected, the last row of the ticker section is identified and therefore we can use the same method to find the closing price for each ticker.

To identify the opening price the code again looks for a change in ticker symbol, however, because we want the first row in the data and not the last row in the data, the code identifies the change in reverse. The code compares the first ticker row with the line above it to identify the change.

Once the opening and closing prices are identified the code can then calculate the price change and the percentage change.

The total stock is calculated by identifying which lines are the same and then adding up each rows total.

Conditional formatting is then applied to the summary table by determining if the yearly change is greater or less than 0. If the change is greater than 0, green is applied and if the change is less than 0, red is applied.

The code then starts the bonus section requirements.
Similarly to the first section, we start by going through each row in the summary table and identifying if each row has a greater stock volume than the next row. If the next row is bigger than the previous, that row becomes the new volume to beat and gets compared to the next one.
The code then applies the same principles to identify the greatest percentage increase and decrease. 

### Results

Once the code has been completed the results for each year appear as per below: 

####  2014

![2014 results](https://user-images.githubusercontent.com/82348616/118746814-3e9fc880-b89c-11eb-8a6c-09307bd5e29c.PNG)

####  2015

![2015 results](https://user-images.githubusercontent.com/82348616/118746828-4495a980-b89c-11eb-939c-7d0da0226e54.PNG)

####  2016

![2016 results](https://user-images.githubusercontent.com/82348616/118746842-4bbcb780-b89c-11eb-8a8b-bedbad8dd693.PNG)

### Files

To run the code please use: _VBA challenge finished code.vb_

Also included:
VBA testing – includes the ‘test data spreadsheet’ and the code used to develop/test.
Screenshots – includes screenshots of the results.


