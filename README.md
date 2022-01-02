# Stock Analysis

## Overview of Project

### Purpose & Background

-The purpose of this module was to use VBA to analyze a large list of stock data to see which ones performed well in the years 2017 & 2018. 
After we got our code running, it was time to refactor it so it would run more efficiently (taking fewer steps, using less memory, etc). 
The end goal of refactoring our original code is so we can compare the run time before and after (we expect the refactored code to run faster).

## Results

-In the original code we used a nested for loop instead of using arrays to store our values (*startingPrice*, *endingPrice*, *totalVolume*). 
The table that is our output was done on each iteration at the end of the first for loop following the nested for loop. 
The following image shows what the code looked like originally:

![green_stocks_code](https://github.com/brand0j/Stock-Analysis/blob/main/Resources/green_stocks_code.PNG)


-With this method we were able to get the table we desired. After doing some simple number formatting and color coding of our output, the table was easily readable. 
At a glance we could tell which stocks did well for 2017 & 2018 (the year is an input value given by the user). 
When you look at the following tables you can see that in 2017 every stock had positive returns except for **TERP**. 
In contrast, 2018 shows that every stock had a negative return with the exception of **ENPH** and **RUN**. 
Another thing to note was that **DQ** had the highest return in 2017 and the lowest return in 2018. 
Accompanying each table is the associated runtime after creating a timer variable (startTime & endTime) so we could track how long our code took to finish executing.

![Analysis_2017](https://github.com/brand0j/Stock-Analysis/blob/main/Resources/Analysis_2017.PNG)
![2017_Original](https://github.com/brand0j/Stock-Analysis/blob/main/Resources/2017_Original.PNG)

![Analysis_2018](https://github.com/brand0j/Stock-Analysis/blob/main/Resources/Analysis_2018.PNG)
![2018_Original](https://github.com/brand0j/Stock-Analysis/blob/main/Resources/2018_Original.PNG)


-Our goal after this point is to refactor our code in hopes that we can cut down on the runtime, thus making it more efficient. 
The main difference between the original code and the refactored code is that we utilize arrays to store the values and output them later. 
This was done by taking the nested for loop and splitting it into two separate loops, one for storing our values (and utilizing a ticker index count), the other for displaying them in excel.


![VBA_Challenge_code](https://github.com/brand0j/Stock-Analysis/blob/main/Resources/VBA_Challenge_code.PNG)

![VBA_Challenge_code2](https://github.com/brand0j/Stock-Analysis/blob/main/Resources/VBA_Challenge_code2.PNG)


-While the output of the table is exactly the same, we can see that the runtime has been significantly reduced:

    
![VBA_Challenge_2017](https://github.com/brand0j/Stock-Analysis/blob/main/Resources/VBA_Challenge_2017.png)
![VBA_Challenge_2018](https://github.com/brand0j/Stock-Analysis/blob/main/Resources/VBA_Challenge_2018.PNG)



## Summary

- What are the advantages and disadvantages of refactoring code?

One of the clearest advantages of refactoring code is that is helped reduce our runtime by a significant amount since we used arrays to store our values to output them later. 
In the initial code we were reading through the table, making calculations, and outputing the results all in one step. 
Removing this nested for loop makes the code easier for someone else to read and would also make it easier to update & maintain in the future. 
One of the main disadvantages of refactoring code is that you are more prone to running into bugs since you are introducing new variables and changing how it functions.

- How do these pros and cons apply to refactoring the original VBA script?

In the original code our runtime was ~0.695s, compared to the refactored code runtime which was only ~0.0743s. The refactored code is approximately 10x faster! 
When considering extremely large data sets it's important to make the code run more efficiently, especially if it is continuously updating (in the context of this project it would be using stock api in real time). 
When refactoring the code for this project it would be easy to run into errors/bugs since we introduced a tickerIndex and chose to use arrays which we needed to use index notation for all of our values. 
This made the process a little tedious double checking that the indexes were correct and the variable names were properly changed throughout the code. 
    

