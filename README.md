# VBA Challenge

## Overview of Project
The purpose of this VBA analysis project was to determine which stocks contained within the provided VBA_Challenge.xlsx file provided the best annual returns in 2017 and 2018. The output of this script will allow novice stock market investors make informed decisions on which "green" stocks are the best to invest in based on historical stocks data.

## Results

### Analysis Results from Refactored Code
The results from the analysis scripts are shown below.
_____

<img width="579" alt="All_Stocks_2017" src="https://user-images.githubusercontent.com/80941606/191327317-7c1a00b9-e8ed-4b5e-8880-601dad333f6d.png">

**Table 1**: This is the output table for the 2017 analysis script. This matches the output table from the original script.

_____

<img width="579" alt="All_Stocks_2018" src="https://user-images.githubusercontent.com/80941606/191327427-904fa5e9-2f0f-415f-8ca7-4555d495338d.png">

**Table 2**: This is the output table for the 2018 analysis script. This matches the output table from the original script.

_____


### Analysis Script Execution Times
The excution times of the original and refactored analysis scripts for the years 2017 and 2018 are shown below. 

_____


<img width="1433" alt="VBA_Challenge_2017_Original" src="https://user-images.githubusercontent.com/80941606/191331267-11dd683c-4610-4b00-acea-5ab529230967.png">

**Image 1**: This is the analysis script's duration message for the original 2017 analysis.

_____


<img width="1433" alt="VBA_Challenge_2018_Original" src="https://user-images.githubusercontent.com/80941606/191331317-da745927-e7df-4b0e-90da-6d529c645982.png">

**Image 2**: This is the analysis script's duration message for the original 2018 analysis.

_____


![VBA_Challenge_2017](https://user-images.githubusercontent.com/80941606/191327686-f5733aae-9fe3-4b78-b63c-331005c250b6.png)

**Image 3**: This is the analysis script's duration message for the new 2017 analysis.

_____


![VBA_Challenge_2018](https://user-images.githubusercontent.com/80941606/191327799-cf8ed010-4a52-4a9a-a5c0-880a10a67d6a.png)

**Image 4**: This is the analysis script's duration message for the new 2018 analysis.

_____

As can be seen in the images above, the refactored analysis script for 2017 was 0.78/0.13 = 6 times faster and the refactored analysis script for 2018 was 0.81/0.18 = 4.5 times faster than the original analysis scripts for 2017 and 2018, respectively. Consequently, the refractored analysis scripts were significantly faster than the original scripts.

### Script Improvements
This improvement in the analysis script's efficieny was achieved by replacing the following code in the original script with code that does not include nested for loops (as seen in the macro in the VBA_Challenge.xlsm file) in the refactored code:

```
'4) Loop through tickers
   For i = 0 to 11
       ticker = tickers(i)
       totalVolume = 0
       '5) loop through rows in the data
       Worksheets("2018").Activate
       For j = 2 to RowCount
           '5a) Get total volume for current ticker
           If Cells(j, 1).Value = ticker Then

               totalVolume = totalVolume + Cells(j, 8).Value

           End If
           '5b) get starting price for current ticker
           If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               startingPrice = Cells(j, 6).Value

           End If

           '5c) get ending price for current ticker
           If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then

               endingPrice = Cells(j, 6).Value

           End If
       Next j
       '6) Output data for current ticker
       Worksheets("All Stocks Analysis").Activate
       Cells(4 + i, 1).Value = ticker
       Cells(4 + i, 2).Value = totalVolume
       Cells(4 + i, 3).Value = endingPrice / startingPrice - 1

   Next i

```

The above original code's refactored version is as follows:

```
''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        'Note: Although one can assume that the tickerIndex starts at 0, I have decided to confirm the tickerIndex with the following conditional statements
        If Cells(i, 1).Value = tickers(0) Then
           'This block of code is expected to always run
            tickerIndex = 0
            tickerVolumes(0) = tickerVolumes(0) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(1) Then
            tickerIndex = 1
            tickerVolumes(1) = tickerVolumes(1) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(2) Then
            tickerIndex = 2
            tickerVolumes(2) = tickerVolumes(2) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(3) Then
            tickerIndex = 3
            tickerVolumes(3) = tickerVolumes(3) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(4) Then
            tickerIndex = 4
            tickerVolumes(4) = tickerVolumes(4) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(5) Then
            tickerIndex = 5
            tickerVolumes(5) = tickerVolumes(5) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(6) Then
            tickerIndex = 6
            tickerVolumes(6) = tickerVolumes(6) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(7) Then
            tickerIndex = 7
            tickerVolumes(7) = tickerVolumes(7) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(8) Then
            tickerIndex = 8
            tickerVolumes(8) = tickerVolumes(8) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(9) Then
            tickerIndex = 9
            tickerVolumes(9) = tickerVolumes(9) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(10) Then
            tickerIndex = 10
            tickerVolumes(10) = tickerVolumes(10) + Cells(i, 8).Value
        
        ElseIf Cells(i, 1).Value = tickers(11) Then
            tickerIndex = 11
            tickerVolumes(11) = tickerVolumes(11) + Cells(i, 8).Value
       End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        '3c) check if the current row is the last row with the selected ticker
         If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
            'The tickerIndex does not need to be incremented due to the conditionals in step 3a)
            'tickerIndex = tickerIndex + 1
        End If
    
    Next i

```

## Summary

### General Advantages and Disadvantages of Refactoring Code
In summary, an advantage of refactoring code is that the code becomes more efficient, which will allow us to process larger datasets in a timely manner. Additionally, refactoring code can also make the code more readable. However, a disadvantage of refactoring code is that the code's efficiency may not improve enough to justify spending time on the refactoring process. Moreover, refactoring code may introduce new problems (i.e. bugs) into the codebase.

### Advantages of Disadvantages of Refactoring the Stocks Analysis Script
In this context, refactoring the original analysis script did signficantly improve the performance of the original analysis script, so refactoring was justified in this case. However, refactoring the code did introduce new bugs into the code which took some time to fix.
