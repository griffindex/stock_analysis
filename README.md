# stock-analysis

## Overview
This is an Excel analysis on 2017-2018 stocks in the Energy sector. The client desired an overview of specific stock tickers that included ticker name, daily volume, and percentage returned. Using Excel Macros, we created seperate sheets from the raw data provided, scripted a button to execute the macro, and included a message box that shows the code's run time.\
### The AllStocksAnalysis macro is structure in the following way:
Initialize an input box and timer variable.\
Initialize an array of all relevant tickers. (These must be manually set with existing tickers in the raw data.)\
Initialize a tickerIndex variable to hold a sum.\
Set the tickerIndex variable to zero using a for loop.\
Set output arrays for tickerVolume, tickerStartingPrices, and TickerEndingPrices\
Use conditionals to increase the sum variable by a value when loopping through the input year's sheet.\
End the loops and include code to conditional format the output sheet.

## Results
The client asked for Total Daily Volume in 2018 for the company DAQU Ticker: DQ.\
The analysis found that this stock's volume declined approximately 62%\
Another analysis that the client requested was over all the stocks in the raw data from the 2017 and 2018 sheets.
These had less volume in 2018 compared to 2017. The refactored code executed faster than the originial code, because loops were used to parse through each array assigned to the variable.
![Image of 2017 Timer Output](/resources/VBA_Challenge_2017.png)
![Image of 2018 Timer Output](/resources/VBA_Challenge_2018.png)

## Summary
The tickerIndex variable is used to access and store data frome each ticker into the output array variables.\
The disadvantage was ensuring that each For loop was closed and that the tickerIndex was referenced properly.\
The advantage is that the code ran more efficiently when refactored.


