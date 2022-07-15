# VBA_Challenge
Challenge 2 working with the Stock Analysis and VBAs to display Color Coordinated Data
# Refactor VBA Code and Measure Performance

## Overview of Project
Using the already given data, I created a VBA that used the 2017 and 2018 stock data to determine how the multiple stocks did over the courses of the year. This data would then be used to make an informed decision on what stocks to invest in. The given data sets contain the open, close, low, high, adjusted high, and the volume in the market each day; this data is then used to determine the yearly stock return as well as the volume of each stock. The VBA creates an input box for the user to input which fiscal year they would like to calculate, once entered, the VBA starts a timer and calculates the returns of the stocks, then when finished, the timer stops and a message box is displayed showing how much time it took for the chosen year to be calculated, as well as color coordinating the returns, red for negative returns and green for positive ones.
### Purpose
The purpose of this project is to take the yearly data of multiple stocks to evaluate which ones are worth investing money into. In order to properly find and display which ones are good to invest in, the return has to be positive. This project was done to make sure that I am able to look over the available data and am able use other data sets appopriately, finding and displyaing important data, and display how the data can be interpreted to make it easier for those I'm coding for.
## Analysis and Challenges

### Analysis of 2017 Stock Information
 I started in the VBA by creating a tickerIndex variable and initializing it to 0, so that it always resets itself whenever the code gets rerun. I then created three arrays, a long called tickerVolumes, and two called tickerStartingPrices and tickerEndingPrices that are Singles, which can have decimals. I then used tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value to add any new volume to the total amount of volume. It then uses code to make sure that the data coming through is in a valid format and to see what row is currently being examined. The tickerIndex is then incremented and through the following functions: Cells(4 + i, 1).Value = tickers(i)
Cells(4 + i, 2).Value = tickerVolumes(i)
Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
we are able to put the acquired values in the given cells on the All Stocks Analysis sheet.
![image](https://github.com/CharlesBootCamp/VBA_Challenge/blob/main/Resources/VBA_Challenge_2017.png)
Ultimately, I took all 12 stocks and used their pieces of data over the fiscal year of 2017 and found that all but one of the stocks, TERP, had gone up in value over the fiscal year.
### Analysis of 2018 Stock Information
 I started in the VBA by creating a tickerIndex variable and initializing it to 0, so that it always resets itself whenever the code gets rerun. I then created three arrays, a long called tickerVolumes, and two called tickerStartingPrices and tickerEndingPrices that are Singles, which can have decimals. I then used tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value to add any new volume to the total amount of volume. It then uses code to make sure that the data coming through is in a valid format and to see what row is currently being examined. The tickerIndex is then incremented and through the following functions: Cells(4 + i, 1).Value = tickers(i)
Cells(4 + i, 2).Value = tickerVolumes(i)
Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
we are able to put the acquired values in the given cells on the All Stocks Analysis sheet.
![image](https://github.com/CharlesBootCamp/VBA_Challenge/blob/main/Resources/VBA_Challenge_2018.png)
Ultimately, I took all 12 stocks and used their pieces of data over the fiscal year of 2018 and found that all but two of the stocks, ENPH and RUN, have gone down in value over the fiscal year.


### Challenges and Difficulties Encountered
One challenge that I experienced was being able to properly the capture the message box that displayed the time that each calculation took and the year being examined, I worked around this by zooming into the message box before taking a screenshot. Another issue I had was I needed to add in an End if statement for the part that said I had to check if the current row is the last row with the selected ticker. A loop then goes through and initilaizes all values in the arrays to 0, to make sure that there is no messes in the data. 
## Results
- What are the advantages and disadvantages of refactoring code?
  Some advantages I learned from refactoring code is that it gives me multiple ways of being able to come across solving the same kind of problem. Another advantage is that the refactoring gives more detail into the nature of the returns and whether or not they are positive or negative. A disadvantage that also comes with this is that it can get confusing and tricky trying to figure out what method is best to complete.
- What are the advantages and disadvantages of the original and refactored code?
The advantage of the original code is that it is simple enough for everyone to understand and do no matter how fresh in language someone is, but an advantage the refactored version can get stuff done incredibly fast. One disadvantage of the original method is that it is ultimately slower than the refactored version, and the disadvantage of the refactored is that it takes more specialization and making sure all the little parts work together.
