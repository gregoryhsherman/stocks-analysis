# stocks-analysis

## Analyzing Stocks

#### Overview

This excersize allowed us to sort through thousands of cells of raw data, identify, quantify, and eventually use said data to show the monetary gains and losses in two fiscal years.  The first part of the excersize, the module, had us target one specific stock from our data and pull the information we needed.  After we took the pertinent data we put it through a script to show just how much money it had made or lost for the year.  After that, we were able to take what we knew for one variable and, through loops, figure out how to apply it to the rest of the variables.  I have attached the vbs script in another readme file that should explain how we got from one code to multiple.

#### Results

The first code that we wrote was good, but after learning more about loops and nested loops, we were able to trim down the run time significantly. By creating arrays and indices to hold values that we defined, we were then able to reference them through our code and make it much more efficient. The pictures linked should prove the superiority of our second iteration:

[2017 First Attempt](https://github.com/gregoryhsherman/stocks-analysis/blob/master/VBA_first_iteration.png)  [2018 First Attempt](https://github.com/gregoryhsherman/stocks-analysis/blob/master/VBA_first_iteration2.png)

As shown, the first attempt is not a bad piece of code - sorting through thousands of cells to define specific parameters.

[2017 Refactored](https://github.com/gregoryhsherman/stocks-analysis/blob/master/VBA_Challenge_2017.png)  [2018 Refactored](https://github.com/gregoryhsherman/stocks-analysis/blob/master/VBA_Challenge_2018.png)

The refactored code, however, is much, much quicker and more efficient. Rather than running loops ad naseum the code referenced our defined instructions to pull the correct data faster.

#### Summary

Overall, the second iteration is vastly superior for our specific example.  With just a few extra lines of code, we were bring the run time down to around twelve percent of the original scripts time.  Overall, the more efficient the script, the better.  Adding arrays and indexes benefited us greatly.
Having seen this for the first time this week, I am not sure I'm qualified to speak on the effectiveness across the board.  For example, if instead of the twelve stocks we were monitoring we were monitoring 3000, adding each specific value in our index would not make sense.  Also, the more loops created within loops seems to drive the run time up greatly, so minimizing the number of nested loops seemed to benefit us.
