# VBA-Challenge
Executive Summary


The market had overall growth in the years 2018 and 2020 while stock prices dropped in the year 2019. 

Introduction


This data contains a years worth of stock information regarding open and closing price, as well as volume sold in a given day. Once completed, this analysis will contain important information regarding the performance of the individual stocks such as the change in a given stocks price over the course of the year, the percent change of the stock, and the total volume traded throughout the year. 


Data Collection and Preparation


This data was provided to me through the course and no preparation or collection was done. 


Data Exploration and Cleaning

I examined the data to understand exactly what was being asked of me and what functionalities my VBA code would need to have. Knowing that the ticker codes ended and would provide a place for a loop to gather all the needed data.  


Data Analysis


For this analysis I looped through the ticker codes on each worksheet until the next code did not equal the one before it. open_Price was taken from the first date a ticker was introduced and close_price was taken from the final row that it appears. Using that information I was able to determine the yearlyChange and percentageChange, as well as determining totalVolume by calculating how many shares of a certain stock were sold in the year.  
According to this analysis, 2020 was the strongest year in terms of year over year stock price change for the entire market. 2018 was the strongest performer in percentage change and also had the highest median stock price increase and -$0.01. Meanwhile 2019 was the only year to show negative overall growth with 52% of companies showing negative growth. 
In the future, I should try to increase the funcionality of my code by attempting to single out the stocks that had the highest and lowest percentage change year over year.


Conclusion


2018 & 2020 both experienced overall positive growth via overall percentage change of all the stocks listed. However, the general performance and percentage change of
stocks from 2019 was quite poor with the average percentage change being negative for the stocks provided. 



