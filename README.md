# Stock-analysis

## Overview of Project

### Objective
The purpose of our analysis was to help a client, Steve, guide his parents in making informed decisions about which green energy company (out of a set of 12) would be the best investment. In order to do so, we use Visual Basic to build macros which will calculate the total daily volume, and percent return, based upon aggregate data of the 12 different green energy companies ([the data set to be used can be found here](Resources/VBA_Challenge.xlsm)). We format the information so Steve is left with a tool which is simple to use and easy to interpret. 

Further, we provide two versions of the VBA macros, one which simply gets the job done, and another which has been optimized for speed. This way, if Steve needs to apply the method we've provided him to extremely large sets of data in the future, he has a means of reducing the time it takes to run the analysis.

>*As a note to the grader, the original "yearAnalysis" subroutine (including a formatting section) is included as module 5 in the .xlsm file, and the refactored code is saved in module 6.*

### Background Information

As mentioned above, we will calculate two parameters in order to analyze the stock performances of the 12 companies:

1. __Total Daily Volume__  
The total daily volume refers to the sum of all the stocks that were traded that day. It can indicate the health of a stock, as higher volumes indicate more interest in that particular company ([read more here](https://www.investorsunderground.com/stock-volume/)). In our analysis, we find total daily volume by simply adding all the volume amounts for a given company.

2. __Percent Return__  
The percent return of a stock indicates net growth or net loss of a stock over a given time period ([read more here](https://finance.zacks.com/stock-market-returns-work-6598.html)). It is valuable in assesing the likelyhood that an investment in a company will yield profit. Here, we find the return by looking at the percent difference between the open value on the first date, and the close value of the last date.

## Results

### Stock Performance

First we want to use our analysis to assess the stock performance of the 12 companies of interest. Below, the results of the calculations for the two years are given (Fig 1 shows 2017 and Fig 2 shows 2018).

![alt text](https://github.com/aamotz001/Stock-analysis/blob/main/Resources/Stocks_2017.png)

__Figure 1: Stock data 2017__

Figure 1 shows that overall, the companies did fairly well, with only the "TERP" ticker showing a negative return (and even so, at roughly 7% it is not catastrophic). As well, the total daily volumes are of large magnitude, assuring us that our results are relevant on a large scale. The company which Steve parent had chosen ("DQ") seems to have performed well but does have a relatively small total daily volume as compared to other companies in this analysis.

![alt text](https://github.com/aamotz001/Stock-analysis/blob/main/Resources/Stocks_2018.png)

__Figure 2: Stock data 2018__

Figure 2 shows that overall, the year 2018 must have been much harder on green energy companies as a whole, with all but two of them showing negative returns. The percent value of the negative returns are severe, especially for the company of interest "DQ" which shows a massive -63%. 

Overall, based on this analysis, we can advise that "DQ" is likely not the best choice to buy. The best bet would be either "ENPH" or "RUN", which are the only two companies with positive returns for both 2017 and 2018. It would be helpful to look at data for more years, and for different types of companies in order to paint a more clear picture. 

### Execution Time

![alt text](https://github.com/aamotz001/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
__Figure 3.a: Stock data 2017 Refactored Runtime__

![alt text](https://github.com/aamotz001/Stock-analysis/blob/main/Resources/VBA_Challenge_OLD_2017.png)
__Figure 3.b: Stock data 2017 Original Runtime__

![alt text](https://github.com/aamotz001/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
__Figure 4.a: Stock data 2018 Refactored Runtime__

![alt text](https://github.com/aamotz001/Stock-analysis/blob/main/Resources/VBA_Challenge_OLD_2018.png)
__Figure 3.b: Stock data 2018 Original Runtime__
## Summary

### What are advantages/disadvantages of refactoring code

### How do these pros and cons apply to refactoring the original VBA script



