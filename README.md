# Overview of Project
The purpose of this analysis was to provide a better look at stocks prices of companies over time. We Looked at the starting price and ending price of 12 different companies and used that data to find the percentage change for the year. Additionally, we found the total volume of trades for the company for the year by adding together all the volumes listed.
This data gives us a idea of how these 12 different companies did for the year.

# Results
Looking at the results we can see a couple of things. In general, green stocks had a much better year in 2017 than in 2018. In 2018 the only well performing stocks were "ENPH" and "RUN". These two stocks also had some of the highest trade volumes. As for "DQ" our target company we can give a reasonable assurance that this is not a good stock to put money into based on our analysis. The code for our analysis looks like
 
    For i = 2 To RowCount
  
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
       
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
            
        If Cells(i + 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
        
            tickerIndex = tickerIndex + 1
        End If
    Next i
   
The "For" loop allow the code to look at every row in the data. The "If" statements select out the start and end that we then assign to the starting price and ending price and we add the volumes in between to get our total volume. We achieved the return volume by dividing the ending price by the starting price and subtracting 1. This looks like `Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1` in our code. Each individual "Ticker" is represented by turning tickerEndingPrices and tickerStartingPrices into array data sets.


Our Code originally looked through the data multiple times slowing the macro down. by building arrays we significantly reduced the time needed to run the program. this is given by the screenshots below. Our pre-optimized code started like this
```
For i = 0 To 11
ticker = tickers(i)
For j = rowStart To rowEnd
```      
As started before this meant the code looked through ALL the data 12 times whereas our new code only needs to look once.

## Pre-Optimization Screenshots
![Preoptimization_2017](https://user-images.githubusercontent.com/96025706/147890179-63acb1b7-fd0d-483b-9602-aaee7654fbc9.png)
![Preoptimization_2018](https://user-images.githubusercontent.com/96025706/147890180-44d142c5-16e5-4c39-9b6a-8e32969b1bdf.png)
## Post-Optimization Screenshots
![VBA_Challenge_2017](https://user-images.githubusercontent.com/96025706/147890190-c561ea91-91dc-4b09-b388-fd7226f0c18b.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/96025706/147890192-3bb5997d-5393-4018-ade4-5baf38736a9b.png)

# Summary
## General
There are advantages and disadvantages to refactoring code. Advantages include the increased productivity gained by making the code run faster for the end user. Additionally refactoring code may also make the code simpler to read and easier for the next coder to work on the code. However, the disadvantages include the amount of time it may take for someone to go through working code to make it better. It may not be worthwhile for a business to pay someone to look over already working code only for it to make a negligible difference in the performance of the code. At the same time, just as the code could become simpler the code may also become more complex and hard to read in the future.
## Specific
For our specific case the advantage is the code is performed faster. However, the disadvantage is that if new tickers are added into our data 5 arrays must be adjusted instead of just 1 for our previous iteration of the code.
