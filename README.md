# Green Energy Stock Analysis
## Overview of Project

For the second project for the UC Davis Data Analytics class, we are using VBA in Excel to do an analysis of stock prices and daily trade volume for green energy stocks.  The goal of the analysis is so that Steve can take our findings and make recommendations to his parents on which green energy stock companies may be better  invest in than the one they chose based solely on its name.  The class project also focuses on the benefits of refactoring code, which we have performed with the goal  to increase execution performance of the VBA script being used to do the analysis.

## Analysis 

### 2017 vs 2018 green energy stock comparison
 We performed an analysis of 12 green energy stock companies, looking at their trade volume and rate of return for each year.  To get this infomation, we used VBA to aggregate the daily volume for each stock from our source data, as well as compare the starting price for the year with the ending price to gain the annual rate of return.  Included below are the total yearly volume and rate of return for each stock for 2017 and 2018.  The goal was to determine if the stock Steve's parents chose, "DQ", was in fact a good investment, and if not what other stock may provide a better profit based on their past performances.  Looking at the results, it seems that "DQ" may not be the best option for Steve's parents to invest.  In 2017 the stock had a high rate of return, 199.4%, but it had a rather low Total Volume of only 35 million.  The following year the volume increased, but the stock price lost 62% of it's value, the biggest loss of all 12 stocks analyzed.  In contrast, another company "ENPH" had a decent daily volume both years (221 million in 2017 and 607 million in 2018) and it returned a profit both years (129.5% in 2017 and 81.9% in 2018), one of only two of the researched companies to return a profit in 2018.  


![2017 Stock Analysis](./resources/2017_stock_analysis.png) 
![2018 Stock Analysis](./resources/2018_stock_analysis.png) 

### VBA Code refactoring
 The second part of our project was to refactor our VBA code in Excel and compare two different coding algorithms to see if we could improve the speed of our code.  Code efficiency is important since we want to be able to run it as the size of our data increases without the process taking exceedingly long.  We therefore did a performance test, of duration it took to complete the analysis, to determine if we could increase the speed of the process through refactoring our code.  In our original code, looped through each stock ticker, and for every ticker we looped through the entire data set to get the information we were looking for.  In our refactored code, we tried only looping through the fulll data set once, rather than once for every ticker in the tickers array, while tracking the data for each ticker based on the ticker value in the current row.  The performance of the analysis improved dramatically as the time it took to execute our code decreased from 0.82 secs down to only 0.14 seconds for 2018 as illustrated below.  


![Code time for 2018](./resources/VBA_Challenge_2018.png) 
![Original Code time for 2018](./resources/VBA_Challenge_2018_original.png) 
![Code time for 2018](./resources/VBA_Challenge_2017.png) 
![Original Code time for 2018](./resources/VBA_Challenge_2017_original.png) 

### Summary
In both versions of our code, the stock prices of the available companies we analyzed received the same results.  However, by changing the algorithms we used to search the data, we were able to reduce the time it took to process the results by a factor of 6.  While neither of these execution times are great, both are under one second, the difference between them will become very noticable as the amount of data increases. Increasing the data by 100 times would have resulted in one version taking 1 min 20 secs, with the other taking just 13 seconds.  Refactoring our code can help us find bugs in the code that we may have missed, it can also help us improve the performance of the code by improving it's efficiency.  In addition, we can add comments we may have missed to increase readability as well.  

There are pros and cons to both versions of the code.  In the original code, it was simpler to write.  I had to keep track of less variables and I could do the calculations for each stock and be done with it.  In the refactored code, I had to keep track of all the stocks until I was done looping through the whole data set.  The advantage that my refactored code had over the original was, of course, that I only had to go through the entire data set one time, not 12 times (or as many times as I had stocks to check).  This led to the code running much faster.  Of course since this is what I was aiming for, the refactored code was the better of the two methods for achieving the same goal, which is the analysis of the stocks.


