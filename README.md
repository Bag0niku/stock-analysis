# Stock-Analysis

## Overview of Project

Steve's Parents' want their newly graduated son to be their financial investment advisor and has asked him to put all of their money into the stock ticker "DQ", a green energy company called DAQO.  Steve is worried that would be to much risk for his parents and that they might need to diversify their portfolio a little. Steve has provided some stock ticker data of the "Green Energy Sector" and asked me to write some vba scripts to help him analyse a stocks performance in the market and compare it against others he has been tracking aswell. 

### Purpose
Make some VBA scripts for Steve to be able to use on a regular baises.

## Analysis 

I have 2 years of data for the same 11 ticker symbols ("AY", "CSIQ", "DQ", "ENPH", "FSLR", "HASI", "JKS", "RUN", "SEDG", "SPWR", "TERP", "VSLR"), each year in it's own worksheet. On the output worksheet use buttons to activate the analysis script or to clear the worksheet. I put a timer in the vba scripts to track the duration of each analysis that was run, for those that are curious about how long it takes.

![2017 Analysis Results](/Resources/VBA_2017.png)
![2018 Analysis Reuslts](/Resources/VBA_2018.png)

##  Results

2017 was a good year, only one ticker finished the year negatively. 2018 was a bad year for most of them, only two of the eleven tickers had a positive year. Comparing the two years together the top 3 stock choices are "RUN", "ENPH" and "SEDG".  The company that Steve's Parents' like ("DQ") did good in 2017 like the rest, but also like the rest did horribly in 2018.


## Refactoring
I started the analysis with the "green_stocks.xlsm" file writing the data to the output worksheet as it was found. I later refactored it to save all the information in 3 arrays before writing to the output worksheet and saved it as "VBA_Challenge.xlsm". Below are comparison times from running the refactored code.

![2017 Analysis Results](/Resources/VBA_Challenge_2017.png)
![2018 Analysis Reuslts](/Resources/VBA_Challenge_2018.png)


Refactoring code is usually done for one of two reasons: 1) to make the code more efficient, or 2) to make it more readable. I do not think i changed the readability of the code, however my refactored code was faster than my original attempt. The one disadvantage of refactoring code I can think of is that you end up doing the opposite, making it less readable and/or more cumbersome.

