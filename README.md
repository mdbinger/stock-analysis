# Overview of the VBA Challenge
## Module 2 of Data Analytics Bootcamp
### Explain the purpose of this analysis.
The purpose of this analysis is to write code in VBA that can take in large amounts of data from the stock market in excel and filter through that data and present us with summarized information regarding the performance of each stock depending on the year. Additionally, we added some components within the code to help with the UI, including formatting the table that the data was presented in, and creating buttons to make running the analysis very easy and accessible for any user from within the excel workbook. Lastly, added some features that help the code run with under less strict parameters, checked how quickly VBA was running the analysis, and we refactored the code substituting in a tickerIndex variable and uses arrays to work with multiple variables at once, which eliminated the need for nested for loop in our code and helped the code run 3 to 4 times as quickly.

# Results of Analysis and Refactored Code
### Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script
### How the Stocks Performed
Overall, 2017 was a much better year for this group of stocks than 2018 was. All stocks except TERP saw an annual increase in value in 2017. However, the near opposite occurred in 2018, as every stock saw an annual decrease in value except RUN and ENPH. ENPH is the only stock that increased in value across both 2017 and 2018 and it was a considerable amount each year!
![2017_Stock_Analysis](https://user-images.githubusercontent.com/96350388/148700800-41bc6290-1bba-44cc-afd9-fae628af92d7.jpg)
![2018_Stock_Analysis](https://user-images.githubusercontent.com/96350388/148700807-126bd20d-48bb-4074-9b12-c99482c571cf.jpg)
### How the Refactored Code Performed
Between the two code scripts used to perform this analysis, the refactored code performed much better. Below are screenshots showing the time it took for the refactored code to run the 2017 and 2018 stock analysis, respectively. 
![VBA_Challenge_2017](https://user-images.githubusercontent.com/96350388/148700836-60920c65-1a46-48dc-90b7-2de7d6723ce0.jpg)
![VBA_Challenge_2018](https://user-images.githubusercontent.com/96350388/148700839-1fcf810f-663c-404e-8d3c-bf028d3e4ff1.jpg)
### How the Original Code Performed
The refactored code results above are both close to 4 times as fast as the original codes time, which are shown below
![2017_Original_Code_Time](https://user-images.githubusercontent.com/96350388/148700848-c11eba82-f7b4-441a-b2d1-af5f0276624d.jpg)
![2018_Original_Code_Time](https://user-images.githubusercontent.com/96350388/148700850-9a03796e-3f2e-4e3a-a421-13de4e527f7e.jpg)
### Differences Between Original and Refactored Code
The major differences in the code which allowed for this more efficient run-time are the use of a nested for loop in the original code, which relies on the first for loop to reset our variable values at the start and output them at the end. The nested for loop searches through all the rows and assigns values to the variables as it goes. 

        '4) Loop through tickers
        For i = 0 To 11
         '5) loop through rows in the data
         ticker = tickers(i)
         totalVolume = 0
        
        Worksheets(yearValue).Activate
        
        For j = 2 To RowCount
        
            '5a) Get total volume for current ticker
            If Cells(j, 1).Value = ticker Then
            totalVolume = totalVolume + Cells(j, 8).Value
                
            End If
        
            '5b) get starting price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j - 1, 1).Value <> ticker Then
            startingPrice = Cells(j, 6).Value
            
            End If
            
            '5c) get ending price for current ticker
            If Cells(j, 1).Value = ticker And Cells(j + 1, 1).Value <> ticker Then
                endingPrice = Cells(j, 6).Value
                
            End If
        
        Next j
        
        '6) Output data for current ticker
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = ticker
        Cells(4 + i, 2).Value = totalVolume
        Cells(4 + i, 3).Value = endingPrice / startingPrice - 1
        
    Next i
        
This is compared to the use arrays functioning with a variable named tickerIndex in the refactored code, which basically allows us to analyze all the tickers in the same loop that they are being assigned values. 
        
        For i = 2 To RowCount
        'Increase volume for current ticker
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(i, 8).Value
        
        'Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i - 1, 1).Value <> tickers(tickerIndex) Then
            tickerStartingPrices(tickerIndex) = Cells(i, 6).Value
        End If
        
        'check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerEndingPrices(tickerIndex) = Cells(i, 6).Value
            
        End If
            
            'Increase the tickerIndex.
        If Cells(i, 1).Value = tickers(tickerIndex) And Cells(i + 1, 1).Value <> tickers(tickerIndex) Then
            tickerIndex = tickerIndex + 1
        
        End If
    
    Next i
In the original code, we are constantly resetting variable values, establishing a new ticker, setting the values for that ticker, then going back starting over again until we run out of tickers. In the refactored code, the ticker changes appropriately as the for loop is assigning data to our variables, so we don’t need to constantly reset out variables because the variables use arrays and their value changes as the ticker changes, so no new ticker will store data over another ticker’s data. 

# Summary of the Impact of Refactored Code
### Advantages of Refactoring Code
There are several potential advantages of refactoring code, including but not limited to: making the code easier to read and/or understand, improving the code’s efficiency, and allowing for the code to take in additional data points that may be added later.
### Disadvantages of Refactoring Code
Typically, refactoring code is done because it is advantageous to do, however, sometimes there might be some unintended consequences. Some of these disadvantages to refactoring code are: it may cause your code to break if done improperly, or the refactored code may be more advanced and less accessible/understandable to beginner coders or people who do not understand code.
### How the pros and cons apply to refactoring our original VBA script
In the case of refactoring our original VBA script, our refactored code used arrays, which is a slightly more advanced coding technique. This hit on a few of the advantages and disadvantages listed above. The use of arrays made our code run much more efficiently and made the code somewhat cleaner to read (if you know what arrays are and how they function). However, if you do not know what arrays are or how they function, the refactored code would become pretty mysterious. Additionally, while refactoring the code it took several attempts to get right, resulting in broken code a number of times. This is frustrating when you originally already had code that worked just fine, and all you were attempting to do was improve it. Ultimately, the refactored code became a better, more efficient program than the original. 
