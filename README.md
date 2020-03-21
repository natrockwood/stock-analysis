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
