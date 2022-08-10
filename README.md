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

### The first image below shows the time it took the refactored code to run for the year 2017:

#### ![VBA_Challenge_2017.png](/ResourcesFinal/VBA_Challenge_2017.png)

### The original code took 1.3398 seconds to run. While a seemingly small change, even saving fractions of a second makes a big difference when running code especially when dealing with large data sets.

### The next image below shows the time it took the refactored code to run for the year 2018:

#### ![VBA_Challenge_2018.png](/ResourcesFinal/VBA_Challenge_2018.png)

### The original code took 1.3633 seconds to run. Again, this is a significant improvement and will run much more efficiently for large data sets. 

## Summary

### Advantages of Refactoring

#### As shown in this analysis, refactoring code has advantages when done correctly. Refactoring saves time and can improve the software to run faster and more efficiently. Refactoring often makes the code easier for others to understand by improving the design of the software. It can also allow for bugs to be found more easily if the code is cleaner and more organized.

### Disadvantages of Refactoring

#### While refactoring has clear advantages there are also disadvantages. Refactoring can be risky, especially when dealing with big applications and can introduce bugs and cause more problems rather than fixing them. If the developer does not fully understand the code refactoring should not be attempted. From a business stand point, refactoring is risky if the deadline is approaching or if the budget does not allow for it. Refactoring should not be done if there is not enough time to test the refactored code before release.

### Advantages and Disadvantages of this VBA Refactoring

#### The advantages of this particular refactoring are the quicker and more efficient run time due to cleaner code. This allows for larger data sets to be analyzed using this code. The code uses only one 'for loop' for the main analysis while the original code used nested 'for loops'; using only one main 'for loop' and storing data in the variables and arrays allows the code to run efficiently. The original code worked and produced the same results, so if the client did not want to use this code for larger sets of data there would have been no point to refactor it; it was running and would have been a waste of time for the software developer. 

