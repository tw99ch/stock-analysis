# Overview of Project

## Explain the purpose of this analysis.

The purpose of this analysis is to helping client to obtain the summarized data for each stock based on total volume and annual return

## Results:
In 2017, all the stocks had positive return except for stock "TERP". "DQ" had the greatest return of 199.4% in 2017 with yearly volume of 35,706,200. There is no correlation between return and volume in 2017.

In 2018, only ENPH and RUN had positive return of 81.9% and 84.0% respectively. The volume for ENPH was 604,473,500 and volume for RUN was 502,757,100. These two stocks outperformed than other stocks with larger volume during the year.  

The execution time for the original code is 1.308594 for 2017 and 1.308594 for 2018.

![](/Run%20time%20for%202017.png)
![](/Run%20time%20for%202018.png)

Whereas the execution time for the refactored script is 0.2578125 for 2017 and 0.2578125 for 2018.

![](/Run%20time%20for%202017%20using%20refactoring.png)
![](/Run%20time%20for%202018%20using%20refactoring.png)

'1a) Create a ticker Index
    Dim tickerIndex As Integer

    '1b) Create three output arrays
    Dim tickerVolumes(12) As Long
    Dim tickerStartingPrices(12) As Single
    Dim tickerEndingPrices(12) As Single
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    ' If the next rowâ€™s ticker doesnâ€™t match, increase the tickerIndex.
    For i = 0 To 11
        tickerVolumes(i) = 0
    Next i
        
    ''2b) Loop over all the rows in the spreadsheet.
    For i = 2 To RowCount
    
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
        tickerIndex = 0
        ElseIf Cells(i, 1).Value <> Cells(i - 1, 1).Value Then
        tickerIndex = tickerIndex + 1
        
        End If
        
        
        '3a) Increase volume for current ticker
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        End If

        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i - 1, 1).Value <> tickers(tickerIndex) And Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
            
            
        'End If
        End If
        '3c) check if the current row is the last row with the selected ticker
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) Then
        tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            

            '3d Increase the tickerIndex.
            
            
        'End If
        End If
    Next i
    
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        tickerIndex = i
        Cells(4 + i, 1).Value = tickers(tickerIndex)
        Cells(4 + i, 2).Value = tickerVolumes(tickerIndex)
        Cells(4 + i, 3).Value = tickerEndingPrices(tickerIndex) / tickerStartingPrices(tickerIndex) - 1
        
        
    Next i
 
 Please refer to the gitHub file for work performed (https://github.com/tw99ch/stock-analysis/blob/master/VBA_Challenge.xlsm)
 

## Summary: In a summary statement, address the following questions.
1. What are the advantages or disadvantages of refactoring code?

    The advantages of the refactoring code is that the script can be executed in shorter time. However, the disadvantages is that the refactoring code is more completed than the original code. If you are not completely understanding the pattern of the script, there is more likely to have an error while writing the refactoring code.

2. How do these pros and cons apply to refactoring the original VBA script?

   The Pros for applying the refactoring code is that any computer can run the script in a much shorter time than the original code. The cons is it takes longer time to write the script because the when the script becomes more complicated, it will also take longer time for me to debug.

