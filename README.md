
Create a script in visual basic that loops through all the stocks for each quarter sheet(4 in total) and outputs the following information:

-The ticker symbol

-Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

-The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

-The total stock volume of the stock. 

-Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 

-Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.


-Make sure to use conditional formatting that will highlight positive change in green and negative change in red.


Code explanation

Basically, I create a loop for going through every tab in the spreadsheet. Then I  creating the Headers like Ticker,Quarterly change. 
Once the header were one I move forward to Copy and paste the first Ticker in the cell H2, because is the first one. After that I create a For loop for going through the first column, skiping the headers, by this i mean first row.Then I defined the first opening price. After that a IF criteria for findig where the tickers are no the same.So by doing this I can take the Opening and closing price for each TICKER.Then  I ade the Quarterly change calculation and Percentage Change calculation base on the information gathered before. Then Appled color conditional formatting to column Quarterly change. After that I added a Loop for calculation Total stock volume. SInce I have all the information in the news column I created another " For loop" for calculation of stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume using the application MIN and MAX. This last portion about Greatest % increase  I used chat GPT for clarifying me how to do it. So I made some ajustment on that code base on my needs.
