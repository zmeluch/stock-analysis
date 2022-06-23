# VBA of Wall Street

## Overview of Project
	
	Steve, a new graduate in finance has a collection of stock data, in an excel workbook, 
	he needs help in analyzing.Steve was interested in gathering data to analyze on green stocks for his parents, his first customers. 
	But the worksheets contain data on all different stocks.To help with this a VBA script was written.

### Purpose
	
	The VBA script helps to quickly automate the sifting and collection of these green stocks data from excel sheet. 
	By searching through the data for specific green stocks and pulling the data for those specific stocks and outputting them 
	into a easy to read formatted sheet. Steve can quickly see what green stocks are doing well or poorly and report that to his parents. 
	After the intial VBA script was written, the code was then refactored.

## Results
	
	In 2017, the green stocks, Steve is analyzing did very well. Only one stock had a negative return. In 2018 however, 
	that was not the case only two stocks had postive returns. Those two stocks though did exceedingly well with over 80% returns. 

	The refactored code, was able to cut the run time on both years' analysis. By adding an index and arrays, 
	we are able to loop through both sheets of data once instead of swtich between the two different years' sheets. 
	This makes the code more efficent and makes it run quicker. As well as makes it easier to understand and work with in the future.

### Stock Analysis for 2017
![2017_Green_Stocks](https://user-images.githubusercontent.com/103155045/175184802-0329ccfb-00be-4792-8c0e-1041ec058a11.png)


### Stock Analysis for 2018
![2018_Green_Stocks](https://user-images.githubusercontent.com/103155045/175184812-00645dde-6db0-49c1-b002-a79b86c613ba.png)

### Original Sub Routine


Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate

    Range("A1").Value = "DAQO (Ticker: DQ)"

    'Create a Header Row

    Cells(3, 1).Value = "Year"

    Cells(3, 2).Value = "Total Daily Volume"

    Cells(3, 3).Value = "Return"

    
    Worksheets("2018").Activate
    
    totalVolume = 0
    
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'Establish the number of row to loop over
    
    rowStart = 2
    
    'rowEnd code taken from https://stackoverflow.com/questions/18088729/row-count-where-data-exists
    
    rowEnd = Cells(Rows.Count, "A").End(xlUp).Row
    
    'loop over all the rows
    
    
    For i = rowStart To rowEnd
    
       
        
        If Cells(i, 1).Value = "DQ" Then
        
         'increase totalVolume if ticker is "DQ"
        totalVolume = totalVolume + Cells(i, 8).Value
        
        End If
        
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        
            startingPrice = Cells(i, 6).Value
            
        End If
        
        If Cells(i + 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then
        
            endingPrice = Cells(i, 6).Value
            
        End If
        
    Next i
    
    
    Worksheets("DQ Analysis").Activate
    
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = (endingPrice / startingPrice) - 1

### Refactored Sub Routine
Sub AllStocksAnalysisRefactored()
    Dim startTime As Single
    Dim endTime  As Single

    yearValue = InputBox("What year would you like to run the analysis on?")

    startTime = Timer
    
    'Format the output sheet on All Stocks Analysis worksheet
    Worksheets("All Stocks Analysis").Activate
    
    Range("A1").Value = "All Stock (" + yearValue + ")"
    
    'Create a header row
    Cells(3, 1).Value = "Ticker"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"

    'Initialize array of all tickers
    Dim tickers(12) As String
    
    tickers(0) = "AY"
    tickers(1) = "CSIQ"
    tickers(2) = "DQ"
    tickers(3) = "ENPH"
    tickers(4) = "FSLR"
    tickers(5) = "HASI"
    tickers(6) = "JKS"
    tickers(7) = "RUN"
    tickers(8) = "SEDG"
    tickers(9) = "SPWR"
    tickers(10) = "TERP"
    tickers(11) = "VSLR"
    
    'Activate data worksheet
    Worksheets(yearValue).Activate
    
    'Get the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    '1a) Create a ticker Index
    Dim tickerIndex As Integer
    
    tickerIndex = 0

    '1b) Create three output arrays
    Dim tickerVolumes(11) As Long
    Dim tickerStartingPrices(11) As Single
    Dim tickerEndingPrices(11) As Single
    
    
    ''2a) Create a for loop to initialize the tickerVolumes to zero.
    For i = 0 To 11
   tickerVolumes(i) = 0
   
        Next i
    ''2b) Loop over all the rows in the spreadsheet.
    For j = 2 To RowCount
    
        '3a) Increase volume for current ticker
         If Cells(j, 1).Value = tickers(tickerIndex) Then
         
        
        tickerVolumes(tickerIndex) = tickerVolumes(tickerIndex) + Cells(j, 8).Value
        
        End If
        
        '3b) Check if the current row is the first row with the selected tickerIndex.
        'If  Then
        If Cells(j - 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
         
         tickerStartingPrices(tickerIndex) = Cells(j, 6).Value
            
            
        'End If
        End If
        '3c) check if the current row is the last row with the selected ticker
         'If the next row’s ticker doesn’t match, increase the tickerIndex.
        'If  Then
            
            If Cells(j + 1, 1).Value <> tickers(tickerIndex) And Cells(j, 1).Value = tickers(tickerIndex) Then
                tickerEndingPrices(tickerIndex) = Cells(j, 6).Value
           
            '3d Increase the tickerIndex.
            
           tickerIndex = tickerIndex + 1
           
            'End If
            
            End If
            
         Next j
   
    '4) Loop through your arrays to output the Ticker, Total Daily Volume, and Return.
    For i = 0 To 11
        
        Worksheets("All Stocks Analysis").Activate
        Cells(4 + i, 1).Value = tickers(i)
        Cells(4 + i, 2).Value = tickerVolumes(i)
        Cells(4 + i, 3).Value = tickerEndingPrices(i) / tickerStartingPrices(i) - 1
        
    Next i
    
    'Formatting
    Worksheets("All Stocks Analysis").Activate
    Range("A3:C3").Font.FontStyle = "Bold"
    Range("A3:C3").Borders(xlEdgeBottom).LineStyle = xlContinuous
    Range("B4:B15").NumberFormat = "#,##0"
    Range("C4:C15").NumberFormat = "0.0%"
    Columns("B").AutoFit

    dataRowStart = 4
    dataRowEnd = 15

    For i = dataRowStart To dataRowEnd
        
        If Cells(i, 3) > 0 Then
            
            Cells(i, 3).Interior.Color = vbGreen
            
        Else
        
            Cells(i, 3).Interior.Color = vbRed
            
        End If
        
    Next i
 
    endTime = Timer
    MsgBox "This code ran in " & (endTime - startTime) & " seconds for the year " & (yearValue)


End Sub

	
### Run times of Original VBA Script and Refactored VBA Script for 2017
![VBA_Challenge_2017](https://user-images.githubusercontent.com/103155045/175184824-adb1f07e-6688-407f-bb6c-a5d0c143a349.png)


### Run times of Original VBA Script and Refactored VBA Script for 2018
![VBA_Challenge_2018](https://user-images.githubusercontent.com/103155045/175184830-4c64c7b7-0b9b-4d60-93ea-82efdebc1fee.png)


## Summary
	
	The advantages of refactoring the code are that the code runs faster. The code also becomes cleaner, 
	easier to read and understand. Making it to work on in the future for oneself and others
	
	In this VBA script the code runs quicker for both data from 2017 and 2018. 
	Also by eliminating Nested loops inside the code becomes easier to read and understand. 
	And eliminates multiply variable which increases the chances of an error.
