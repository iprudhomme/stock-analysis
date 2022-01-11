# Green Energy Stock Analysis
## Overview of Project

As the second project for the UC Davis Data Analytics class, we are using VBA in Excel to do an analysis of stock prices for Green Energy stocks so that Steve can make recommendations to his parents on which companies to invest in.  The project also focuses on the benefits of refactoring code to increase performance.

## Analysis 

An analysis of the stock volumes and yearly returns for the 12 green energy stocks seems to indicate that DQ may not be the best option for Steve's parents to invest.  In 2017 the stock had a high rate of return, 199.4%, but it had a rather now Total Volume of only 35 million.  The following year the volume increased, but the stock price lost 62% of it's value.  In contrast, ENPH had a decent daily volume both years (221 million in 2017 and 607 million in 2018) and it returned a profit both years (129.5% in 2017 and 81.9% in 2018), one of only two of the researched companies to return a profit in 2018.  

In addition, we did a performance test of the analysis to determine if we could increase the speed of the process by refactoring our code.  Using a VBA timer to measure the run time of the calculations.  We refactored the code so that we only looped through the full data set once, rather than once for every ticker in the tickers array.  The performance of the analysis improved as the duration decreased from 0.82 secs down to only 0.14 seconds for 2018.

![Code time for 2018](./resources/VBA_Challenge_2018.png) 
![Original Code time for 2018](./resources/VBA_Challenge_2018_original.png) 

For 2017, the code times decreased as well from 0.84 down to 0.14 as well.

![Code time for 2018](./resources/VBA_Challenge_2017.png) 
![Original Code time for 2018](./resources/VBA_Challenge_2017_original.png) 


### Summary
In both versions of the code to analyze the stock prices of the available companies we received the same results.  However, we reduced the time it took to process the results by a factor of 6.  Refactoring our code can help us to find bugs we may have missed, it can help us improve the performance of the code by improving it's efficiency.  We can add comments we may have missed to increase readability as well.  

There are pros and cons to both versions of the code.  In the original code, it was simpler to loop through the code since the outer loop was representing the ticker index. 


