# stock-analysis
An analysis of Stock data from 2017 and 2018 using Excel

## Overview
To see if there is a relation between total volume of a stock's trades and a stock's yearly return. A Macro was created using VBA to find the stocks. The Macro was refatcored to find the results faster.

## Results
### Stock Analysis
![Stock ananlysis Results](https://github.com/aneou/stock-analysis/blob/main/Resources/Stock_Results.png)
###
As scene in the chart above while the stock with the most trades each year did have postivive returns, the stock with the most return relative to its starting price in 2017 was DQ which had the lowest trades of the stocks listed. Also SEDG and VSLR had less trades in 2017 than 2018 but had greater returns in 2017. These show that there is no correalation between a stock's trade volume and return.
### Code Refactoring
![Original Runtime](https://github.com/aneou/stock-analysis/blob/main/Resources/AllStockAnalysis_Runtime.png)
![Refactored Runtime](https://github.com/aneou/stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
###
The orginal VBA script took around three quarters of a second while the refactored script took about a seventh of that time at around a tenth of a second.

## Summary 
While this project did not have useful results on the relation between Trade Volume and stock returns, it did provide valuable experience on refactoring code. Refactoring code requires an understanding of what your code is doing and how it is doing it. It then needs to find the deficiencies with the code and improve upon them. For example in the original script used in this project the code looped over the entire list for each ticker, however this was unnecssary as the trades had been sorted by tickers first. 
