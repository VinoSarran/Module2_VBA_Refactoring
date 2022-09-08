# Refactoring with VBA: You built it.  Now build it better.

## Overview of Project:

### Many would argue that the key to successful investing is diversification.  There is value in evaluating companies on the frontier of “Green” technology and innovation but history has shown us time and time again that there is no surefire strategy to finding success in the stock market.  Using Excel and VBA, I have built a tool to aggregate a year’s worth of key metrics for green stocks thus allowing potential investors to gauge which stocks are worth exploring based on prior performance.    However, evaluating these types of companies in isolation may be risky.  At a minimum, additional stocks should be added to the analysis to provide a necessary benchmark to measure green stocks against.  Perhaps, stocks tied to other forms of energy.  And as previously mentioned, smart investors should diversify.  Beyond green stocks and benchmark stocks, potential investors should use this Excel/VBA tool to evaluate the entire market.  The VBA-based tool created to evaluate green stocks will inevitably be stressed as demands on it increase.  In anticipation of that, additional effort was made to create efficiencies.  The initial script was evaluated and restructured to leverage arrays and non-nested loops within VBA.  This will consume less processing memory and improve runtime even as more and more stocks are appended to the source data.  
 <br>

## Results:

- The original version of this tool leveraged fairly common practices of establishing an array as a list then cycling through each cell within the specified range. The original design read and evaluated each row flipping back and forth from the source data to the output table. A loop has created for each ticker and result was printed after each loop:  
            
            For i = 0 To 11
            ticker = tickers(i)
            
  Using total volume as example, we can see the straightforward logic to evaluate the initial row compared to the which ticker in the index was being evaluated and      complete the task once the ticker changed:
        
                   For i = 0 To 11
            ticker = tickers(i)
            totalvolume = 0
'5.Loop through rows in the data.
                 Worksheets(yearValue).Activate
                 For j = 2 To RowCount

'   5a.Find the total volume for the current ticker.
                    If Cells(j, 1).Value = ticker Then
                    'increase totalVolume by the value in the current row
                    totalvolume = totalvolume + Cells(j, 8).Value
                    End If
'   5b.Find the starting price for the current ticker.
                    If Cells(j - 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    startingPrice = Cells(j, 6).Value
                    End If
'   5c.Find the ending price for the current ticker.
                    If Cells(j + 1, 1).Value <> ticker And Cells(j, 1).Value = ticker Then
                    endingPrice = Cells(j, 6).Value
                    End If
                Next j
'6.Output the data for the current ticker.
            Worksheets("All Stocks Analysis").Activate
            Cells(4 + i, 1).Value = ticker
            Cells(4 + i, 2).Value = totalvolume
            Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
            
        Next i

    This original methodology produced a runtime of:
    
    <br>
![model](https://github.com/VinoSarran/Module2_VBA_Refactoring/blob/8e88873fc9afbfe73676b5d4c0bd16003f6b8a72/VBA_Challenge_2018%20(2).png?raw=true)

- The refactored version of this tool leverages use of arrays to copy the source data in to memory to avoid flipping back and forth from the source data to the output table.  The array only stores the necessary values outlined.  Each subsequent loop references the array in memory rather than reading the source data directly using the code below:   
             For j = 2 To RowCount
                    tickerVolumes(tickerindex) = tickerVolumes(tickerindex) + Cells(j, 8).Value
 
                    If Cells(j - 1, 1).Value <> tickers(tickerindex) Then
                    tickerStartingPrices(tickerindex) = Cells(j, 6).Value
                    End If
 
                    If Cells(j + 1, 1).Value <> tickers(tickerindex) Then
                    tickerEndingPrices(tickerindex) = Cells(j, 6).Value
                    
                    tickerindex = tickerindex + 1
              Next
              
   Once all the required calculations are made within memory, the version of the code simply copies the array output to the output cells on the worksheet.  
   
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

The differences between the original analysis code and the refactored illustrate the versality available when building tools for end users.  Often times, we can get to the same end by many different means.  But to be successful, we must think about how a tool will be used and how its use could evolve.  One of the advantages of the code in its original state is its ease of use and the ability of minimally trained coders to alter the logic.  However, it did not run as effectively as possible.  Refracturing does often lead to efficiency but a challenge is ensuring the code is still performing the same task and ensuring the logic has not changed.  Having the ability to compare both versions was essential.       
 
Ensuring the clients are able to evaluate as much data as possible was crucial.  However, the VBA script is not fool proof.  This code relies heavily on the source data being properly sorted.  Tickers out of sequence will report incorrect information.  The refactoring exercise does not address this fault and end users may not be aware of this issue. 
