# Refactoring with VBA: You built it.  Now build it better.

## Overview of Project:

### Many would argue that the key to successful investing is diversification.  There is value in evaluating companies on the frontier of “Green” technology and innovation but history has shown us time and time again that there is no surefire strategy to finding success in the stock market.  Using Excel and VBA, I have built a tool to aggregate a year’s worth of key metrics for green stocks thus allowing potential investors to gauge which stocks are worth exploring based on prior performance.    However, evaluating these types of companies in isolation may be risky.  At a minimum, additional stocks should be added to the analysis to provide a necessary benchmark to measure green stocks against.  Perhaps, stocks tied to other forms of energy.  And as previously mentioned, smart investors should diversify.  Beyond green stocks and benchmark stocks, potential investors should use this Excel/VBA tool to evaluate the entire market.  The VBA-based tool created to evaluate green stocks will inevitably be stressed as demands on it increase.  In anticipation of that, additional effort was made to create efficiencies.  The initial script was evaluated and restructured to leverage arrays and non-nested loops within VBA.  This will consume less processing memory and improve runtime even as more and more stocks are appended to the source data.  
 <br>

## Results:

The original version of this tool leveraged fairly common practices of establishing an array as a list then cycling through each cell within the specified range. The original design gather all information for an individual stock and presented final calcualtions in output before moving on to the next ticker.  This is essentially  flipping back and forth from the source data to the output table.  A loop was created for each ticker and result was printed after each loop:  
            
            For i = 0 To 11
            ticker = tickers(i)
            
We can see the straightforward logic to evaluate the initial row and comparing it to the specific ticker in the index that was being evaluated.  This also causes the need for mutliple IF-ENDIF statements:
        
            For i = 0 To 11
            ticker = tickers(i)
            totalvolume = 0
 
                 Worksheets(yearValue).Activate
                 For j = 2 To RowCount

                    If Cells(j, 1).Value = ticker Then
                    'increase totalVolume by the value in the current row
                    totalvolume = totalvolume + Cells(j, 8).Value
                    End If
 
                    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    startingPrice = Cells(j, 6).Value
                    End If
 
                    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    endingPrice = Cells(j, 6).Value
                    End If
                    
                Next j
                
Finally we see that the output worksheet is activiated and output is created. This means the loop will have to activate the source data once again as the loop advances. 

            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalvolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
            
        Next i

This original methodology produced a runtime of:
    
![model](https://github.com/VinoSarran/Module2_VBA_Refactoring/blob/8e88873fc9afbfe73676b5d4c0bd16003f6b8a72/VBA_Challenge_2018%20(2).png?raw=true)

Let's discuss how the refactored version differs from the above.  The refactored code leverages arrays to copy the source data into memory to avoid flipping back and forth from the source data to the output table.  The array only stores the necessary values outlined. Each subsequent condition references the array in memory rather than reading the source data directly.  As a result, I was able to remove nested loops and allow the logic to be processed in sequence.  And advantage of using an array for each ticker is that total volumes can be summed without IF-ENDIF:   
            
            For j = 2 To RowCount
                    tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(j, 8).Value
 
                    If Cells(j - 1, 1).Value <> tickers(tickerindex) Then
                    tickerStartingPrices(tickerindex) = Cells(j, 6).Value
                    End If
 
                    If Cells(j + 1, 1).Value <> tickers(tickerindex) Then
                    tickerEndingPrices(tickerindex) = Cells(j, 6).Value
                    
                    tickerindex = tickerindex + 1
              Next
              
Once all the required calculations are made within memory, this version of the code simply copies the array output to the output cells on the worksheet using its own loop.  
   
    For i = 0 To 11
            
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = tickers(i)
            Cells(4 + i, 2).Value = tickerVolumes(i)
            Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
            
        Next i
   
   
   
   This refactored version has a runtime of:
       <br>
 ![alt text](https://github.com/VinoSarran/Module2_VBA_Refactoring/blob/main/VBA_Challenge_2018.png?raw=true)


 <br>

## Summary: 

Now clients are able to step beyond the original scope of evaluating just Green stocks.  Benchmark stocks or even the entire market could be evaluated along side those Green Stocks.  

The differences between the original analysis code and the refactored illustrate the versatility available when building tools for end users.  Often times, we can get to the same end by many different means.  But to be successful, we must think about how a tool will be used and how its use could evolve.  One of the advantages of the code in its original state is its ease of use and the ability of minimally trained coders to alter the logic.  However, it did not run as effectively as possible.  Refactoring does often lead to efficiency but a challenge is ensuring the code is still performing the same task and ensuring the logic has not changed.  Having the ability to compare both versions was essential.       
 
Ensuring the clients are able to evaluate as much data as possible was crucial.  In addition to efficiency, an advantage of the refactored code is the ability to transform the data as you build your array.  This could be utilized in future refactoring exercises.  However, this VBA script is not fool proof.  This code relies heavily on the source data being properly sorted.  Tickers out of sequence or unsorted will cause the macros to fail.  The refactoring exercise does not address this fault and end users may not be aware of this issue. In addition, the code is not completely ready to perform the task of analyzing the entire stock market.  Tickers need to be called out explicitly which will be time consuming and undo a lot of the efficiencies created.
