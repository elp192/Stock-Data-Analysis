# Analysis of Stocks Data 
## Project Overview
### Background
The Excel sheets provide data related to green energy stocks during the 2017 and 2018 years. By analyzing the dataset, it can be determined whether the investment in a particular stock is worthy or not.<br>
Each sheet related to 2017 or 2018 consists of about 3000 rows and 8 columns. In total, information about 12 different stocks is available. Columns include some attributes associated with the stock ticker, date, opening and closing price of the stock, lowest and highest daily price of the stock, and daily volume of stock.
### Purpose
In this project, we will help Steve, who wants to analyze and understand a trend of the green energy markets for his parents. His parents are only interested in investing their money in DQ, a new energy corporation that makes silicon wafers for solar panels. However, Steve wants to analyze other energy stocks besides DQ to make sure they get the most advantage from their investment.<br>
The data are analyzed using Visual Basic for Applications (VBA). Several conditionals and nested loop are utilized. Also, a user-friendly interface is created, enabling a user to run the functions by only clicking the buttons. Finally, the code is refactored to improve its logic and efficiency.
## Results
### Analysis of different stocks data for 2017 and 2018
The "All Stocks Analysis" sheet with three columns named Ticker, Total Daily Volume, and Return is created to analyze the data. Figure 1 represents the total daily volume and percentage of yearly return for all the green energy stocks in 2017 and 2018.<br>
From Figure 1 (left), it can be observed that in 2017, the percentage of yearly return for all the stocks except one of them (i.e., TERP) is positive and is varied in the range of 5%-200%. However, we can interpret that this trend is different in 2018 (Figure 1, right). In this year the percentage of the return is negative for most of the stocks (10 out of 12). We can conclude that for the majority of stocks, 2017 was the better year for investment compared to 2018.<br>
Steve‘s parents were interested in DQ stock. Results show that for this stock, although the volume of daily trading in 2018 is increased considerably compared to 2017, the rate of yearly return is significantly decreased from 199.4% to -62.6%. So, it can be concluded that the performance of DQ stock is unsatisfactory, and it is an unstable and unpredictable stock for investment.<br>
<p img align="center" width="100%">
   <img width="231" alt="AllStocks_2017" src="https://user-images.githubusercontent.com/85843401/124953838-7cd77f80-dfe3-11eb-9546-b49c350fcae2.png">
   <img width="232" alt="AllStocks_2018" src="https://user-images.githubusercontent.com/85843401/124953855-81039d00-dfe3-11eb-98bf-99b57c8c0eb0.png"><figcaption>Figure 1: Representing the total daily volume, and  yearly return of 12 stocks in left) 2017 and right) 2018.</figcaption></figure/> 
<p align="center">
</p>

### Analysis of refactoring code
Two modules with the same outputs (Figure 1) are created in [green_stocks.xlsm]. In the following, the difference between these modules and the time of execution are discussed.<br>
The main difference between the two modules is that in Module 2, Module 1's code is refactored in which the nested loop is substituted by the For loops. Figures 3 and 4 (left) represent part of the code before refactoring and after refactoring, respectively. As shown in Figure 3 (left), in the outer For loop, two different worksheets are activated in each iteration, which can increase the execution time. However, Figure 4 (left) shows the refactored code that the inner loop is eliminated, and an array named ticker volume and a variable called ticker index are defined. In addition, to store the outputs in a worksheet, another separated For loop is created. Figures 3 and 4 (right) show that after refactoring code, the execution time is decreased from 0.62 seconds to 0.11 seconds.<br>

<p img align="center" width="100%">
   <img width="270" alt="Before refactoring code" src="https://user-images.githubusercontent.com/85843401/124977658-bf5a8580-dffe-11eb-97db-ee0b541af394.png">
   <img width="240" alt="Execution time_before refactoring" src="https://user-images.githubusercontent.com/85843401/124977681-c6819380-dffe-11eb-8218-43cd8b9ab8e5.png"> <figcaption>Figure 3: Module 1 (before refactoring code). Left) Part of the code. Right) Execution time.</figcaption></figure>
</p> 

<p img align="center" width="100%">
   <img width="270" alt="After refactoring code" src="https://user-images.githubusercontent.com/85843401/124977671-c386a300-dffe-11eb-8286-b3b252c355fe.png">
   <img width="240" alt="Execution time_after refactoring" src="https://user-images.githubusercontent.com/85843401/124977697-ca151a80-dffe-11eb-9428-9b4e810d1d61.png">
  <figcaption>Figure 4: Module 2 (after refactoring code). Left) Part of the code. Right) Execution time.</figcaption></figure>
</p> 

## Summary
**The advantages and disadvantages of refactoring code**
Code refactoring is an important process in computer programming to reorganize the code that its behavior remains the same. the advantages and disadvantages of refactoring code, in general, are as follows:<br>
:+1: **Advantages**:<br>
:white_small_square: Refactoring code helps to decrease the execution time and makes the code as efficient as possible.<br>
:white_small_square: Refactoring code improves the readability of code, and therefore, it is easier to follow and understand codes.<br>
:white_small_square: As code is cleaner and the logic of the code is improved, debugging and adding new functionality to code for future purpose is easier.<br>

:-1: **Disadvantages**:<br>
:white_small_square: Refactoring code is a time-consuming process. In other words, it does not guarantee how much time is needed to get a desirable outcome.<br>
:white_small_square: This process may influence the outputs. However, the main aim of refactoring code is to improve the code without manipulating the software's functionality.<br>
### The advantages and disadvantages of the original and refactored VBA script 
In this project, Module 1 (before refactoring code) took about 0.6 seconds to run; however, in Module 2 the run time is decreased by about 82%. In addition, by eliminating the nested For loop, the code is easier to read, and it is not needed to activate and switch between different worksheets in each iteration. <br>
We made a trade-off between gaining advantages of refactoring code and spending time to clean the original code, which gave us the same result as the original one. However, as the dataset is small, the run time for the original code is reasonable, and for the larger dataset, refactoring can be more beneficial.
