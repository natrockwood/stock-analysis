# Stocks Analysis
## Using VBA
The Visual Basic Application or more commonly known as VBA helps in the analysis of large data sets that Excel formulas might not handle.
The most basic VBA progam one can write returns a Message Box that says "Hello World!"
```vba
Sub MacroCheck()
  Dim testMessage As String
  testMessage = "Hello World!"
  MsgBox (testMessage)
End Sub
```
## DQ Analysis
Tried making column headers in the Excel file using VBA codes
```vba
Sub DQAnalysis()
	Worksheets("DQ Analysis").Activate 
	Range("A1").Value = "DAQO(Ticker: DQ)"
	Range("A3").Value = "Year"
	Range("B3").Value = "Total Daily Volume"
	Range("C3").Value = "Return"
End Sub
```
This can also be programmed as:
```vba
Sub DQAnalysis()
	Worksheets("DQ Analysis").Activate
	Cells(3, 1).Value = "Year"
	Cells(3, 2).Value = "Total Daily Volume"
	Cells(3, 3).Value = "Return"
End Sub
```
"Worksheets("DQ Analysis").Activate" tells Excel to activate the DQ Analysis worksheet and put our analysis there
## Using For Loops and Conditionals
In this exercise I was trying to compute for the Total Daily Volume in 2018 if the Ticker was "DQ" \
To do this, I first had to activate the 2018 Worksheet for the Macro to run on.\
I also mades sure my header row in the "DQ Analysis" sheet was all set up.\
Since there are 3013 rows in the 2018 sheet, it would be inefficient to find the X number of "DQ"s there were in that sheet and add it up. In Excel, you can use the formula: 
```excel
=SUMIFS($H:$H,$A:$A,"DQ")

Column H = Volume
Column A = Ticker Values
"DQ" = Value of Volumes we need to sum up
```
In VBA, we can program this by using the below:
```vba
Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    ' Tells Excel to activate the worksheet and put our analysis there
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate

    rowStart = 2
    rowEnd = 3013
    totalVolume = 0
    
    'For i = 1 To 8
        'MsgBox (Cells(1, i))
        
    For i = rowStart To rowEnd
        'There are 3,013 rows in the 2018 worksheet.
        'increase totalVolume if ticker is "DQ"
        If Cells(i, 1).Value = "DQ" Then
            totalVolume = totalVolume + Cells(i, 8).Value
        End If
    Next i
    
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume

End Sub
```
This code returns the sum of  107,873,900, which the Sum of the Total Daily Volume in 2018 for DQ.
## Getting DQ's Yearly Return for 2018
Steve wanted to calculate how well (or how bad) DQ performed in 2018. This is done by  calculating the yearly return which is the +/- % in price from the beginning to end of year.\
From our data, we need the Starting Price and Ending Price.
#### Calculating the Starting and Ending Prices
To calculate for these variables, these were determined by the cell values in the ticker which is "DQ" \
The outcome of these variables were stored as these dimensions:
```vba
Dim startingPrice As Double
Dim endingPrice As Double
```
At the end of the code, we just calculated the Return as the (Ending Price / Starting Price) - 1, since it's a percentage. \
The outcome turned out to be a -63% Return, which is pretty bad, and Steve is most likely to recommend a better stock to invest in to his parents. 

## Full DQ Analysis Code
```vba
Sub DQAnalysis()

    Worksheets("DQ Analysis").Activate
    ' Tells Excel to activate the worksheet and put our analysis there
    
    Range("A1").Value = "DAQO (Ticker: DQ)"
    
    'Create a header row
    Cells(3, 1).Value = "Year"
    Cells(3, 2).Value = "Total Daily Volume"
    Cells(3, 3).Value = "Return"
    
    Worksheets("2018").Activate

    'set initial volume to 0 and startingPrice variable as decimal values
    totalVolume = 0
    Dim startingPrice As Double
    Dim endingPrice As Double
    
    'find the number of rows to loop over
    RowCount = Cells(Rows.Count, "A").End(xlUp).Row
    
    'loop over all the rows
    For i = 2 To RowCount
    
        If Cells(i, 1).Value = "DQ" Then
            'increase totalVolume by the value in the current row
            totalVolume = totalVolume + Cells(i, 8).Value
            
        End If
        
        'checking if the current row’s ticker is DQ and checking if the previous row’s ticker is not DQ
        '(i-1) is the previous row; (i,1) current row
        If Cells(i - 1, 1).Value <> "DQ" And Cells(i, 1).Value = "DQ" Then

            'set starting price with the price data in the 6th column
            startingPrice = Cells(i, 6).Value

        End If
        
        '(i+1) is the next row; (i,1) current row
        If Cells(i, 1).Value = "DQ" And Cells(i + 1, 1).Value <> "DQ" Then

            endingPrice = Cells(i, 6).Value

        End If
        
    Next i
    
    'With the starting and ending prices stored, we can now add a line to our output to show the yearly return for DQ\
    Worksheets("DQ Analysis").Activate
    Cells(4, 1).Value = 2018
    Cells(4, 2).Value = totalVolume
    Cells(4, 3).Value = endingPrice / startingPrice - 1

End Sub
```
