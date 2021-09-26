# VBA Refactoring Green Stock Analysis 

## Overview and Purpose

Steve, a financial analyst, has pulled a large set of Green Company stock and wants a quick way to run and analyze the data for each stocks total daily volume (how many times a day a stock is traded) and the annual return on investment (the stock price change from the beginning of the year to the end of the year). Steve asked for our assistance in updating or refactoring the Visual Basic Application (VBA) code from our original analysis, of the Green Stock dataset, to make it more efficient. 

## Analysis and Results

### VBA Code Analysis With Nested For Loop

In our original analysis of the Green Stock dataset our VBA code was written with a nested For loop, as shown in the image below.

![ Gree_Stock_ For_Loops]( https://github.com/lmacera/Stock-Analysis/blob/main/Resources%202/Gree_Stock_%20For_Loops.PNG )

With the nested For loop our VBA’s first loop read through the ticker array to initialize the stock tickers total daily volume to zero. The second loop would then loop through all stock tickers in column A of the Green Stock data set to increase each tickers array for *totalvolume*, *tickerstartingprices*, and *tickerendingprices*. Once the second, or nested For loop, was finished running through the data set the first For loop would move onto the next ticker, or iterator, to run the second For look again. The information collected, when reading through these loops, was collected using if statements nested in the For loops and stored using the arrays of *tickervolumes*, *tickerstartingprices*, and *tickerendingprices* as show in the below image. 

![Gree_Stock_ IF_Statements]( https://github.com/lmacera/Stock-Analysis/blob/main/Resources%202/Gree_Stock_%20IF_Statements.PNG )

### VBA Code Analysis With Indexing

In refactoring the code our team created the same arrays used in the original analysis, *ticker*, *tickervolume*, *tickerstartingprices*, and *tickerendingprices*; however, instead of referencing the ticker array with a For loop we referenced and increased the ticker array with an index, as shown in the image below.

![ VBA_Challenge_Index]( https://github.com/lmacera/Stock-Analysis/blob/main/Resources%202/VBA_Challenge_Index.PNG )

Then **one** For loop was created to loop through our index and **one** For loop was created to run through column A of the Green Stock dataset. Our first Loop would initialize the *tickervolume* back to zero for the new ticker in the index and our second loop would increase each tickers array for *totalvolume*, *tickerstartingprices*, and *tickerendingprices*. This essentially created a parallel processing system in which each individual ticker had their corresponding array for *totalvolume*, *tickerstartingprices*, and *tickerendingprices* collected at the same time. Like the nested For loop analysis, the information collected when reading through these loops was collected using if statements nested in the For loop and stored using the arrays of *tickervolumes*, *tickerstartingprices*, and *tickerendingprices*, as show in the below image.

![VBA_Challenge_IF Statements]( https://github.com/lmacera/Stock-Analysis/blob/main/Resources%202/VBA_Challenge_IF%20Statements.PNG )

###Results

From our refactoring we were able to speed up the processing time between our original nested For loop and our Indexing analyses. As show in the below images our run times for the nested For loop for 2017 and 2018 was 1.086 seconds and 1.055 seconds respectively.

![Green_Stock_Analysis_ 2017 Run Time]( https://github.com/lmacera/Stock-Analysis/blob/main/Resources%202/Green_Stock_Analysis_%202017%20Run%20Time.PNG )

![Green_Stock_Analysis_ 2018 Run Time]( https://github.com/lmacera/Stock-Analysis/blob/main/Resources%202/Green_Stock_Analysis_%202018%20Run%20Time.PNG )

For our Indexing analysis we saw a decrease in processing times with processing times of 0.24 seconds for 2017 data and 0.48 for 2018 data.

![ VBA_Challenge_ 2017]( https://github.com/lmacera/Stock-Analysis/blob/main/Resources/VBA_Challenge_%202017.PNG )

![ VBA_Challenge_ 2018]( https://github.com/lmacera/Stock-Analysis/blob/main/Resources/VBA_Challenge_Stock_%202018.PNG )

As for the analysis of stock, our data set produced the same information showing regardless of which code we ran. At a high level our analysis shows the majority of the stocks with higher daily volume count in 2017 compared to 2018. Stocks that traded particularly well in each year were ENPH, RUN, SPWR, and FSLR, as show in the below images.

![ VBA_Challenge_ 2017]( https://github.com/lmacera/Stock-Analysis/blob/main/Resources%202/VBA_Challenge_2017.PNG )

![ VBA_Challenge_ 2018]( https://github.com/lmacera/Stock-Analysis/blob/main/Resources%202/VBA_Challenge_2018.PNG )

From these images we can see 2017 was the best year for annual return on investment with only the TERP stock showing a negative return. 2018 however, most stocks had a negative return on investment except for ENPH and RUN.

##Summary

### Pros of Refactoring

There are many pros to refactoring code but from our analysis above we can see the most common benefit of refactoring is the efficiency in processing time. Additionally, refactoring code can help programmers find efficiencies in pattern recognition, in which logical errors will be more easily recognizable, and helps the code to be more “readable” to others.

###Cons of Refactoring

A downside to refactoring code is that it is often time consuming for a process that is already performing the job that is needed. Additionally, it could create more bugs in the code and make the code less and or unstable compared to the original code.

### Pros and Cons of Our Analysis

With regards to our analysis, we saw both the pros and cons of refactoring. Through refactoring our code, we saw that our processing time to produce our analysis was cut in half and our code became cleaner and more “readable”. However, the downside to our refactoring was that it was time consuming for a process that was already running at an efficient pace and producing the same results.

