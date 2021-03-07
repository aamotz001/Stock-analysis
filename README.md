# Stock-analysis

## Overview of Project

### Objective
The purpose of our analysis was to help a client, Steve, guide his parents in making informed decisions about which green energy company (out of a set of 12) would be the best investment. In order to do so, we use Visual Basic to build macros which will calculate the total daily volume, and percent return, based upon aggregate data of the 12 different green energy companies ([the data set to be used can be found here](Resources/VBA_Challenge.xlsm)). We format the information so Steve is left with a tool which is simple to use and easy to interpret. 

Further, we provide two versions of the VBA macros, one which simply gets the job done, and another which has been optimized for speed. This way, if Steve needs to apply the method we've provided him to extremely large sets of data in the future, he has a means of reducing the time it takes to run the analysis.

>*As a note to the grader, the original "yearAnalysis" subroutine (including a formatting section) is included as module 5 in the VBA_Challenge.xlsm file, and the refactored code is saved in module 6.*

### Background Information

As mentioned above, we will calculate two parameters in order to analyze the stock performances of the 12 companies:

1. __Total Daily Volume__  
The total daily volume refers to the sum of all the stocks that were traded that day. It can indicate the health of a stock, as higher volumes indicate more interest in that particular company ([read more here](https://www.investorsunderground.com/stock-volume/)). In our analysis, we find total daily volume by simply adding all the volume amounts for a given company.

2. __Percent Return__  
The percent return of a stock indicates net growth or net loss of a stock over a given time period ([read more here](https://finance.zacks.com/stock-market-returns-work-6598.html)). It is valuable in assessing the likelihood that an investment in a company will yield profit. Here, we find the return by looking at the percent difference between the open value on the first date, and the close value of the last date.

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

Another way in which we used our analysis was to inspect how different styles of writing VBA code can affect the time it takes for the code to run. To do so, we wrote two different codes which accomplished the same task. The first code, called yearValueAnalysis() relies on a nested for loop to first set a ticker of interest (outer for loop) then calculate the relevant parameters by spanning the entire data set (inner for loop). The next analysis was a *refactored* version of the first subroutine which used only one for loop and depended upon an indexing variable called tickerIndex to determine which part of the code to include in the analysis. In both cases, the timing of each run was determined within the VBA code. The results are compared in Figs 3 & 4 (Figure 3 shows the run times for 2017, and Figure 4 shows it for 2018) and the "a" panel shows the original code versus the refactored code in panel "b". 

We can see from the results that the refactored code is significantly faster. Obviously, the second routine which only uses one for loop and only scans the data set once per iteration is much more efficient, however, it does rely upon the data being grouped according to the ticker name. 

![alt text](https://github.com/aamotz001/Stock-analysis/blob/main/Resources/VBA_Challenge_2017.png)
__Figure 3.a: Stock data 2017 Refactored Runtime__

![alt text](https://github.com/aamotz001/Stock-analysis/blob/main/Resources/VBA_Challenge_OLD_2017.png)
__Figure 3.b: Stock data 2017 Original Runtime__

![alt text](https://github.com/aamotz001/Stock-analysis/blob/main/Resources/VBA_Challenge_2018.png)
__Figure 4.a: Stock data 2018 Refactored Runtime__

![alt text](https://github.com/aamotz001/Stock-analysis/blob/main/Resources/VBA_Challenge_OLD_2018.png)
__Figure 4.b: Stock data 2018 Original Runtime__
## Summary

### What are advantages/disadvantages of refactoring code

The obvious advantage of refactoring a code is that it requires much less time to finish the calculations, compared to the disadvantage of the original code being slower. This type of update would become crucial as data sets become larger and larger, or the calculations performed become more and more complicated. However, the disadvantage of refactoring the code may be that it complicates the structure of the VBA file and makes it less robust (it is more specialized to a particular type of notebook format, or data set for example). As well it may be less intuitive for an outsider to look at the refactored code and determine what is happening VS the original code, which is clearer because of the nested for loops.

### How do these pros and cons apply to refactoring the original VBA script

Again, the obvious advantage of refactoring a code is that it requires much less time to finish the calculations. However due to the scale of this worksheet, the difference in time is rather negligible (less than a second difference between the two modalities), and so it may  not be a huge advantage to use the refactored code. If we did get a notebook which looked at say, 500 companies rather than just 12, the advantage of the refactored code would become very obvious. The disadvantage of refactoring the code is that it became less robust, and now depends on all the information being grouped by the ticker name. If we were to apply the code to a new spreadsheet that was not sorted in such a way, we would not get the right answer without first spending time formatting the notebook to match what the refactored code expects - it may be easier to just stick with the original code instead of spending time reformatting the notebook.
