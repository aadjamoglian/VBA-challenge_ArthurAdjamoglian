# The VBA of Wall Street

## Summary

You are well on your way to becoming a programmer and Excel master! This projeact use VBA scripting to analyze real stock market data.

### Files

* [Data](Multiple_year_stock_data_ArthurAdjamoglian.xlsm) - Used this data to run scripts and generate reports.

### Stock market analyst

![stock Market](Images/stockmarket.jpg)

## Features

### Yearly Change, Percent Change, Total Stock Volume, and Formatting

* The script loops through all the stocks for each year for each run and takes the following information.

  * The ticker symbol.

  * Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The percent change from opening price at the beginning of a given year to the closing price at the end of that year.

  * The total stock volume of the stock.

* Conditional formatting will highlight positive change in green and negative change in red.

* The result looks as follows:

![moderate_solution](Images/moderate_solution.png)

### Greatest Percent Increase, Decrease, and Total Volume

* The scipt returns the stock with the "Greatest % increase", "Greatest % Decrease" and "Greatest total volume". The result looks as follows:

![hard_solution](Images/hard_solution.png)

