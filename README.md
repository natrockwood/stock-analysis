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
