# Stock-Analysis
This is an analysis of Green Energy stocks using Visual Basic for the purposes of advising our friend Steve and his parents on investing. This analysis looks at trading dates, initial trading values, final trading values, and trading volumes to assess annual returns for 2017 and 2018.

## Project Overview
The stock market can be overwhelming, making it difficult to determine the best investments. This project utilized several VBA coding tools such as arrays, loops, and conditionals to analyze stock information and creates reports of the annual findings. Developer buttons were included to make it easy for any user to run the code, with a simple interactive input box and color-coded results to call out the gains and losses on annual returns. With these tools Steve and his parents can make better investing decisions.

### Layout
The workbook includes 4 tabs:
- Two data tabs containing 2017 and 2018 stock market data (respectively)
- An "All Stocks Analysis" tab
- A "DQ Analysis" tab

The data tabs contain Green Energy stock:
- Tickers
- Trading dates
- Opening stock prices
- Highest daily stock prices
- Lowest daily stock prices
- Closing (and adjusted closing) stock prices
- Stock volume traded each day

The "All Stocks Analysis" tab contains two developer buttons, one for running the VBA code that asks the user for the year they want to assess and creates a table of results by stock ticker, and one for clearing the worksheet. Similarly, the "DQ Analysis" tab contains similar developer buttons, but instead of looking at all Green Energy stocks it analyzes only DQ stocks. 

## Results
### Beginning the Analysis
The DQ Analysis was created first in order to look at the results of Steve's parents' initial investment in DQ. This looked specifically at 2018 data. 

![image](https://user-images.githubusercontent.com/105830645/173010037-dbc30005-a8ca-4297-a3fd-2cc2a84bd22a.png)

### Building and Expanding on the VBA Code
A VBA subroutine was created to analyze all the Green Energy stocks rather than just the DQ stocks, creating a table of results. The code was further expanded to allow for analysis of both 2017 and 2018, depending on user choice. Upon selecting the "Run Analysis for All Stocks" button the user enters the year to analyze (2017 or 2018). The results are shown below:

![image](https://user-images.githubusercontent.com/105830645/173010501-58763533-e1b5-4d8e-8d11-d8c923ce81c6.png)![image](https://user-images.githubusercontent.com/105830645/173010575-1a866d72-aedc-4e36-88c1-7dd7f7d7db0a.png)


### Comparing the Initial and Refactored Code
For this challenge the initial code created for the Subroutine "yearValueAnalysis" (left) was refactored (used as a basis for and reconfigured) for the Subroutine "AllStocksAnalysisRefactored" (right). The code is shown side-by-side below:

![image](https://user-images.githubusercontent.com/105830645/173015633-4f93d479-dd0e-4fe1-ad8c-b276c41fc88c.png)
![image](https://user-images.githubusercontent.com/105830645/173015664-c177e956-3a1b-4613-afa9-f7ff59b5bf63.png)
![image](https://user-images.githubusercontent.com/105830645/173015679-54c6e12d-e523-49aa-b2a9-9666998626ca.png)

The refactored code shown here uses three output arrays rather than one, which are total volume, total starting price, and total ending price. A tickerIndex was also used throughout the subroutine, allowing it to run more efficiently by eliminating a loop. The time comparison is shown below:

![image](https://user-images.githubusercontent.com/105830645/173016356-7171b0dd-0056-4055-87c9-b33644610b73.png)

The code now runs much faster with the same results!

### Reading the Results
The final tables show a marked decrease on returns for Green Energy stocks in 2018 compared to 2017, with an additional decrease in the overall volume of stocks traded. While Steve's parents didn't lose all their initial investment due to the pronounced return on DQ stocks in 2017, the trend is troubling. By comparing these tables, we see that ENPH has returns for both years, outperforming RUN over the two-year period (the only other stock with returns for both years). 

## Summary
### Pros and Cons of Refactoring Code
When building initial code, the intent is to meet all requirements of the necessary output. This can be done piece-by-piece, carefully looking at each step. Once the initial code is created and tested, refactoring can have some advantages. Focusing on simplification and streamlining steps can reduce the time and resources the code takes to run. There are also lessons learned when refactoring code that can be used in for future projects. 

There are some drawbacks to refactoring. It is a time investment to review and reconfigure the initial code. Errors may be introduced that need to be resolved. For code that takes up few resources with fast run time it may not be worth the investment in time and energy if it already works well for the purpose it was designed. Coder bandwidth should be weighed against resources needed to run code, frequency of use, and the longevity of the project.

### Pros and Cons of Refactoring the Stock-Analysis Code
The original code used for the yearValueAnalysis subroutine included nested loops, essentially requiring the initial program to go through the data top-to-bottom repeatedly. This is eliminated in the refactored code by including three output arrays and a tickerIndex (as noted above). The subroutine ran much faster. For the volume of stock information in the input this is not significant, but for large volumes of stock information using the refactored code would have a great advantage.
