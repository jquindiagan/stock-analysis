# Stock Analysis

## Overview of Project

### Purpose
The purpose of this project was to:
    1. analyze stock market data for 2017 and 2018 using VBA and 
    2. practice refactoring the VBA code.

## Results

### Stock Analysis - 2017 vs 2018
As depicted in the screenshots below, the results show that stocks overall had higher returns in 2017 than they did in 2018. In 2018, nearly all stocks shrunk except for ENPH and RUN. Even then, ENPH had a higher return in 2017. RUN is the only stock that had higher returns in 2018 than it did in 2017.

<img src="Resources/Results_2017.png" width=30% height =30%>  <img src="Resources/Results_2018.png" width=30% height =30%> 

### Refactoring the VBA Code
The original codes both had execution times of about 0.67 seconds. After refactoring the code, there was a slight difference in execution time by about 0.07 seconds.

<b>Original Code</b>

<img src="Resources/VBA_Original_2017.png" width=30% height =30%>  <img src="Resources/VBA_Original_2018.png" width=30% height =30%>

<b>Refactored Code</b>

<img src="Resources/VBA_Challenge_2017.png" width=30% height =30%>  <img src="Resources/VBA_Challenge_2018.png" width=30% height =30%>

## Summary

### Refactoring: Advantages and Disadvantages
Refactoring code comes with the advantage of having a more efficient and organized code. However, there is also the disadvantage of potentially causing errors in the code. For this specific code, I experienced the latter with the "expected array" error. It took some time to find a solution that would allow the code to run.