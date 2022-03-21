## stock-analysis
VBA Module repository

## Requirements
- Microsoft Excel 2016 or greater

### Version list
- Version 0.001.  No work done yet, just enabled VBA tools in Excel.
- Version 0.002.  A file is recieved and backed up and macros are enabled and the file is backed up again.
- Version 0.003.  Adds the first Work in progress update to the file containing the hello world test.
- Version 0.004.  Work in progress version 2.  Added some code to format cells in Excel.  Just row headers.  Baby steps.
- Version 0.005-0.009.  Updates not documented individually due to time constraints.  Added data seeking and formating functions.  Minor bug fixes.
- Version 1.000 First release version, user interface buttons and year query added.
- Version 1.001 Minor engine update and bug fix patch, performance monitoring added.
- Verison 1.002 Performance update for potential scope expansion.

# Overview

- This project is designed to automate analysis of stock data from an Excel worksheet and present it in a user friendly and easily readable breakdown.  Current version is compatable with Stock tickers "AY," "CSIQ," "DQ," "ENPH," "FSLR," "HASI," "JKS," "RUN," "SEDG," "SPWR," "TERP," and "VSLR" with future plans for expansion.

- The original code presented in version 0.009 was presented to the client for initial needs.  Afterward, the client expressed interest in expanding the project beyond the current scope.  Performance monitoring and user experience improvements were made in preparation for this.

# Results
## Analysis
-  ![image 1](Resources\2017secondrefractor.png)
 ![image 2](Resources\2018secondrefractor.png)

- Between the end 2017 and 2018 we can see a decline in the overall performance of the various green energy stocks.  However, the long term performance of seems to be on the increase when taken into the context of the performance between the end of 2016 and 2017.  Of note, while the trend of DQ was downward in 2018 investors from before 2017 would still have seen an overall growth.  Other stocks of note were ENF which has managed to continue growth through the trend, as well as RUN which has had unprecidented growth in 2018.

## Performance of Code.

### Initial Conditions.
- On initial release in version 0.009 loops were used to check each line from the input to check for each ticker symbol from an array in turn.  While this completed the task in just under a second, as shown in the performance screen shot below, the client expressed interest in running either a larger sample of the stock market or the entire market itself at a later time.  This would be many times the current 12 supported symbols, and performance would decline under such a condition.

![image 3](resources/2017firstrefractor.png)


### Solution
- The code was refractored to replace a nested for loop that incrimented a tracking variable through all tickers. This was replaced with a single for loop that incrimented through the ticker array sequentially when the last line for a particular symbol was located by a conditional statment.

```
'initial code
'loop through the tickers
    For c = 0 To 11
        Worksheets(yearValue).Activate 'go to the sheet with the data in it.
        checkTicker = tickers(c)
        totalVolume = 0
        
            For i = rowStart To rowEnd 'creates a for loop that goes through each populated row in 2018
    
                If Cells(i, 1).Value = checkTicker And Cells(i - 1, 1).Value <> checkTicker Then 'checks for the first DQ listing
                    startingPrice = Cells(i, 6).Value 'gets the starting value for the company we're looking for.
                End If
        
                If Cells(i, 1).Value = checkTicker And Cells(i + 1, 1).Value <> checkTicker Then 'checks for the last DQ listing
                    endingPrice = Cells(i, 6).Value 'gets the ending value for the company we're looking for.
                End If
        
                If Cells(i, 1).Value = checkTicker Then  'grabs all the values with a ticker value of DQ in column a
                    totalVolume = totalVolume + Cells(i, 8).Value 'adds up the total value of the traded values in column h
                End If

            Next i

```
```
'code after refractoring and variable reassignment to specifications.

    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        '3a) Increase volume for current ticker
        'For tickerIndex = 0 To 11  'uses ticker index to check the current row against all tickers, uncomment if second code is unstable
            If Cells(i, 1).Value = tickers(tickerIndex) Then  'grabs which ticker this row matches
                tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value 'updates the tickerVolume
            End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
            If Cells(i, 1).Value <> Cells(i - 1, 1).Value And Cells(i, 1) = tickers(tickerIndex) Then
                tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            End If
            
        'End If
        
        '3c) check if the current row is the last row with the selected ticker
         'If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
        'If  Then
            If Cells(i, 1).Value <> Cells(i + 1, 1).Value And Cells(i, 1) = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
                tickerIndex = tickerIndex + 1 'increases the ticker index AFTER reading it's last line. Ment to save cycles at the cost of stability.  Comment out and replace with the tickerIndex for loop if this is unstable
            End If
            

            '3d Increase the tickerIndex.
        'Next tickerIndex 'to use the for loop method.  use if indexing at the last line is unstable.
            
        'End If
    
    Next i
    
```
# Conclusions
## observations
- Refractoring of code decreased the output time of the macro significantly, however this comes at the cost of input error checking.
- As long as the input from the user is formatted properly, the code very quickly and efficiently gives the requested data.  This formatting can be done using basic Excel functions: sorting first by column B then by column A in standard A-Z order.
- Input format is currently strict and output may become unstable if current formatting is not followed.  See worksheet 2017 or 2018 for an example of the required formating.
- All tickers MUST be present within the listed dataset in column A in alphabetical order for the macro to function properly.
- All data MUST be in cronological order per ticker in column B in order for the macro to function properly.
