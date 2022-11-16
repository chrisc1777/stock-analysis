# VBA Challenge

## Overview of Project
Steve originally wanted an Excel VBA code script to analyze a dozen rows of data for DQ's stock performance. After learning that DQ was not a good stock to purchase, Steve wanted to start a new project called "All Stocks Analysis" to analyze the entire stock market from 2017 and 2018. The same piece of code was used in the new project but with the exception of a few variable updates. The code was successful but the execution performance was slow for running thousands of data rows instead of a dozen. 

The purpose of this challenge is to refactor the Excel VBA code script of "All stocks Analysis" in order to shorten the time process of performing the code. 


## Results
The original code executed the "All Stock Analysis" in 0.55 seconds for 2017 and 0.54 seconds for 2018. 

![VBA_Original_2017](https://user-images.githubusercontent.com/106359564/202310473-1b927a31-80d0-45ac-b4f7-692eb6425be3.png)
![VBA_orignal_2018](https://user-images.githubusercontent.com/106359564/202310490-b58eb6a7-9d5c-473b-9238-c0094bcf505d.png)

The refactored code was able to perform in 0.39 seconds for both 2017 and 2018. The chart analysis shows that 2017 had better stock returns with 11 positive increases while 2018.

![VBA_Challenge_2017](https://user-images.githubusercontent.com/106359564/202311009-4ed4a385-ba93-4d89-aae1-f7db556255fa.png)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/106359564/202311022-526f6c4c-8562-48ba-a5d5-41b2d2704c04.png)


The refactored code used a new variable "tickerIndex" and three new output arrays to loop over thousands of data rows.

![VBA_Challenge_Script](https://user-images.githubusercontent.com/106359564/202310557-90be945f-3385-4a93-ae70-a9966f256572.png)


## Summary
Refactoring the VBA code gave advantages in faster execution time and making the variables well-defined to perform the analysis. The disadvantages of refactoring the code included writing a longer script and experience with syntax errors to make the code function.

Although refactoring took more effort in adding variables and making statement changes for the loop to make the code work, it was an overall improvement. The function of the code was superior in handling more data and being more detailed with its variables than the original VBA script.
