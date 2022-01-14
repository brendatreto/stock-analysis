# VBA of Wall Street

## Overview of Project
### Purpose
Steve wants to analyze the behavior of the Stock Market for the years 2017 and 2018. We started by analyzing one stock, the DAQO, but we realized that this process could and should be improved to analyze an entire dataset. 
With this in mind, and with the help of VBA tools we expanded the analysis to include the entire stock market. 
The results were satisfactory, but as it is often said, there is always room for improvement, and this is where this project comes in handy. Refactoring this code could help it run faster and accomplish the same tasks in a more efficient manner. 

### Background
During the last couple of weeks, we were building a VBA script to run through the dataset. In the process, we were learning about loops, conditionals, variables, formatting and other VBA features. However, although the code we wrote to analyze the stocks was working, it was not at its best. For instance the elapsed time to run the analysis was around 1.4 seconds

![2017_runtime](https://user-images.githubusercontent.com/22451540/149431742-09074fe5-363f-40dc-8997-ab25e859dc49.png)

![2018_runtime](https://user-images.githubusercontent.com/22451540/149431751-d8dc73f7-0879-4855-a666-fa0e1a07e377.png)

## Results
### Stock performance
We can tell that this analysis is working by looking at the "All Stocks Analysis" worksheet. Instead of going over hundreds of numbers, we can tell, simply by the color formatting the overall performance of the stocks. Even without diving deep into the data it is evident that 2017 was a better year for green energy stocks, compared to 2018.

This excercise comes in handy when we want to make information available for everyone, even those who have no background experience with the stock market, just like Steve's parents. 
### Refactoring the code
1. We kept the first part of the code as it was. Initialize variables to count the elapsed time of the analysis after asking for the user's input. We also created a header row to organize our information.
2. The next step stayed the same as well, initializing an array for all tickers, activating the data worksheet depending on the user's input and counting the rows to loop over.
3. For this step we began with the process of refactoring by creating a ticker Index that started at zero, with the purpose of standardizing the information before the loop began. 
4. The next step (1b), and to my opinion, the key step in speeding the process, was creating three output arrays, saving us valuable seconds in the analysis.

![Step 4](https://user-images.githubusercontent.com/22451540/149433732-1656c568-a80f-4928-be58-a59a4d8e9c6f.PNG)

5. Once the arrays were declared, we could start with our first for loops (2a) and initialize the volume of each ticker (previously stated in the arrays) in zero.
6. The next step followe pretty much the same structure as the starting code. We first used conditionals to make sure that the rows we were analyzing corresponded to the first and last pieces of information of that ticker, and used that information to increase the volume of each index.
7. We used a different for loop to build the output of our analysis, and again were able to save valuable seconds by referencing the previously created arrays.

![Step 7](https://user-images.githubusercontent.com/22451540/149434384-18c62c8d-1ad2-42d4-8ba5-b1189cb29ad9.PNG)

8. We kept the same code to give format to the results. A key step in the visualization part of the process. We also needed to stop the timer and print the result to make sure that our purpose was achieved.

