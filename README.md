# "Green Energy" Stock Analysis

## Project Overview
The goal of this project is to analyze "Green Energy" stocks using historical stock data. In this analysis, we are looking at 12 different stocks and comparing two specific data points: yearly volume, and yearly return. By calculating these values, we are able to determine which stocks best fit our clients portfolio. We are also anaylizing data from both 2017 and 2018 to get a better understanding of how the stocks have performed on a longer time horizon. 

## Analysis
By writing a VBA script to analyize the data from our 12 sample stocks, we are able to quickly and accurately calculate the two metrics we are using to compare the stocks. As seen in the screenshots below, we can now make simple inferences from our results such as: How has the stock performed over the last couple of years? How liquid is the stock? 

In order to achieve this, we pulled the historical trading data from all 12 stocks for 2017 and 2018. Using ```for loops``` and ```variables``` we are able to write a simple mathmatical function that automatically calculates the values we are after. In the second iteration of the code, we refactored the individual ```for loops``` to a single ```for loop``` and used arrays to simplify the code and speed up the macro's ability to calculate our values. 
```
        For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        
                tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(i, 8).Value
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
            
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
         
                tickerStartingPrices(tickerindex) = Cells(i, 6).Value
            
            End If
        
        '3c) check if the current row is the last row with the selected ticker
             'If the next rows ticker doesnst match, increase the tickerIndex.
            
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
         
                tickerEndingPrices(tickerindex) = Cells(i, 6).Value
         
         '3d Increase the tickerIndex.
            
                tickerindex = tickerindex + 1
            
            End If
```

Follwing the refactoring of the code, we were able to achieve much faster calculation times, as seen in the screenshots below.

### Results
Refactoring a code makes it easier to comprehend as well as increase the process speed (ideally). This is imporant especially in programs that get used often and repetitively. However, by changing the code it is possible to introduce new bugs, and debugging can be a challenge; especially with long and complex code. 

While our original code was able to calculate the correct results, we were able to simplify the code by refactoring it to a single ```for loop``` thus increasing the speed at which the function performed the calculations needed. This also makes it easier to understand and more adaptable for future use. However, by refactoring the code we did have to spend a significant amount of time re-writing the code and debugging it. It is possible that the amount of time spent refactoring the code may have ended up being longer than the time saved by changing it (Depending on how often the program will be used).
