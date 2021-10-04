# stock-analysis

# VBA of Wallstreet

## Overview of Project

### Performing analysis on Green Energy stock options to uncover trends

## Results

### Analysis of stock performance between 2017 and 2018
When analyzing the stock performance it's clear that 2017 was a more successful year than 2018.  In 2017 only one company had a negative return, in 2018 on the other hand, ten companies had a negative return.  Now aside from that year just being more successful, what I found very interesting was how often one of the top five companies in one year were also part of the bottom five companies in the other year.  TERP, RUN, AY, DQ, JKS, and FSLR were all some of the highest performers one year, and the worst the next.  There's actually only two companies who stayed in the bottom five performers across both years and one that stayed in the top five.  With such drastic results from year to year I would highly recommend looking into a data set spanning across a more significant date range to fully understand which options are consistently the best bet.  With the data we are given I would suggest the client invest in ENPH as a primary company and SEDG or VSLR as secondary choices.  
<img width="200" alt="All Stocks 2017 Chart" src="https://user-images.githubusercontent.com/90050622/135935580-3704c2ee-487a-4514-b5c9-c6137bd5af9e.PNG"><img width="200" alt="All Stocks 2018 Chart" src="https://user-images.githubusercontent.com/90050622/135935594-f78720e9-b785-402c-b04c-5fbe32ed523e.PNG">

By adding additional variables to my refactored code, I was able to write short, simplifed *if* statements, such as the one below. 
```
 If Cells(i - 1, 1).Value <> tickers(tickerindex) Then
  tickerStartingPrices(tickerindex) = Cells(i, 6).Value
```
    

## Summary

### Advantages and disadvantages or refactoring code
While there are many advantages to refactoring code, one of the most prominent is the improved readability of the code.  Having easy to read code has many benefits.  As the code author, it is significantly easier to troubleshoot or debug problems if you have more efficient code in front of you.  This is also important if you plan to share your work out to others, who will need to interpret it as well.  Often times when refactoring has not taken place, code smell can be detected.  This could include bad patterns, duplicate code, or just unneccessarily long code.  Refactoring is a great way to improve the quality of the code into something more efficient and effective.  However, with that said, there are some disadvantages to refactoring code.  Refactoring code takes time, and with some projects more time is not an option.  Especially in a business environment, management often doesn't see a problem with clunky code as long as it works, so the concept of refactoring seems like a waste.  Refactoring code also opens the door to potential new bugs.  Just because you've written something that worked once doesn't mean that when you rewrite it it will run smoothly, in fact, it could end up causing more harm than good.  If a certain code lives in a workbook that you plan to use for a long time and add and modify to, refactoring is a great investment, however, if it's a one and done project, it might not be worth the extra work.  


#### How these pros and cons aply to refactoring the original VBA script?
I ran into a lot of issues with my original code in this VBA script.  It was so long, and there were so many moving parts, that when I would add additional steps, or try to debug a problem, I would actually end up with additional problems.  This code was greatly in need of some refactoring.  In it's original state it was long and redundant, and hard to keep track of.  This doesn't mean that refactoring was an easy solution, as I did still run into bugs that I had to correct.  Refactoring meant using more advanced VBA skills, some of which were new to me.  While the original code was much longer, it was written in much more basic code.  There was a clear advantage to refactoring the code in regards to run time.  While this wasn't a huge data set, you could still see very clearly that the refactored code ran faster than the original code.  This would be a crucial factor to keep in mind if you were working with a large dataset. 

Below, we can see the run times of original code for both 2017 and 2018 on the left and run times for refactored code on the right. 

<img width="233" alt="VBA_Challenge 2017 Original code" src="https://user-images.githubusercontent.com/90050622/135935948-4a35858a-9784-4485-9580-9ef5eb47b770.PNG">![VBA_Challenge_2017_Refactored Code](https://user-images.githubusercontent.com/90050622/135935969-a7918489-b469-4624-bb9f-6be48202c84f.png)


<img width="239" alt="VBA_Challenge 2018 Original Code" src="https://user-images.githubusercontent.com/90050622/135936029-81144306-909e-4ac6-bdff-4ea22f8f25a1.PNG">![VBA_Challenge_2018_Refactored Code](https://user-images.githubusercontent.com/90050622/135936032-d06d1b1c-24ef-46ce-933c-fc0f2c63d3dc.png)


