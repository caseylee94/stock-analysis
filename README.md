# Stock Analysis for Steve

## Overview of Project

### Steve is a client who wants a code to analyze the stock market results for different years to do research on which stocks would be a good investment for his parents. I created a code that would run over the data he had provided, which was analyzing a dozen different stocks. Now with this new knowledge, Steve wants to expand his research and analyze thousands of stocks to become even more informed for his parents' investment. The original code I created would not work well for a data set of this size, so the purpose of this project is to refactor the existing code to be more efficient and able to analyze thousands of different stocks quickly.

## Results

### I refactored the code to run more efficiently by creating additional variables, output arrays, and formatting them into the code to allow for only one 'for loop' to need to run to analyze the results. An additional 'for loop' was then created to output the data from the arrays. The original code, while achieving the same results, did not output the results as quickly. These additional variables and arrays allow the code to have somewhere to store data and use the variables to run analysis, which becomes particularly important when analyzing larger sets of data. The original code and the refactored code output the same results, and using the VBA timer function we can see the difference in time it takes for each code to run. 

### Code and Timer Images

#### The output arrays were created using this code:

#####         Dim tickerVolumes(tickersCount) As Long
#####         Dim tickerStartingPrices(tickersCount) As Single
#####         Dim tickerEndingPrices(tickersCount) As Single

#### The ticker index was created with this code:

#####         Dim tickerIndex As Integer
#####         tickerIndex = 0

##### ![VBA_Challenge_2017.png](/ResourcesFinal/VBA_Challenge_2017.png)

### This is the time it took for the refactored code to run for the year 2017; the original code took 1.3398 seconds to run. 

###
![VBA_Challenge_2018.png](/ResourcesFinal/VBA_Challenge_2018.png)
