# VBA-challenge

This VBA sub takes daily stock data sorted by companies' stock market tickers and outputs a summary for each ticker in the data set.

For the first loop, the sub takes the total number of rows used in the sheet as the upper limit. Every row stores and adds the stock volume, which is reset when either if statement is triggered.

For each row with a ticker, the sub checks the previous row to see if the current row is the first occurrence of a ticker. If so, it stores the opening stock price. 

If the previous row has the same ticker, the sub checks the next row to see if the current row is the last occurrence of a ticker. If so, it stores the closing stock price, displays the ticker, and calculates and displays the total price change, percent change, and volume. It then changes the price change fill color if it is positive (green) or negative (red).

For the second loop, the sub takes the number of rows filled in the previous loop as the upper limit. It then looks for the maximum percent change, minimum percent change, and maximum total volume and outputs those values and corresponding tickers.
