' use for each ws

'Create a script that will loop through all the stocks for one year and output the following information.
    'The ticker symbol.
    'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
    'The total stock volume of the stock.

'Create Variables - ticker as sring, rest as double (large numbers)
Sub StockData()

Dim Ticker As String
Dim Yearly_Change As Double
Dim Percent_Change As Double
Dim open_price As Double
Dim Total_Stock As Double
Dim Close_Price As Double
Total_Stock = 0

'Know location of each ticker symbol
Dim SummaryTable As Integer
SummaryTable = 2


'create summary table by location
Cells(1, 11) = "Ticker"
Cells(1, 12) = "Yearly Change"
Cells(1, 13) = "Percent Change"
Cells(1, 14) = "Total Stock Volume"
Columns("M:M").Select
Selection.Style = "Percent"
Selection.NumberFormat = "0,00%"
Columns("L:L").Select
Selection.NumberFormat = "General"

open_price = Cells(2, 3).Value

'Locate last row with formula & variable 
LastRow = Cells(Rows.Count, 1).End(xlUp).Row


'Loop through all rows....Go through each variable to loop & name location in the summary table 

For i = 2 To LastRow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    Ticker = Cells(i, 1).Value
    
    Total_Stock = Total_Stock + Cells(i, 7).Value
    
    Range("K" & SummaryTable).Value = Ticker
    
    Range("N" & SummaryTable).Value = Total_Stock
    
    SummaryTable = SummaryTable + 1
    
    Total_Stock = 0
    
    Close_Price = Cells(i, 6).Value
    Yearly_Change = Close_Price - open_price
    
    Range("L" & SummaryTable - 1).Value = Yearly_Change
    
    
    If open_price > 0 Then
    Percent_Change = (Close_Price / open_price) - 1
    open_price = Cells(i + 1, 3)
    
    Else
    Percent_Change = 0
    
    End If
    
    Range("M" & SummaryTable - 1).Value = Percent_Change
    
Else

Total_Stock = Total_Stock + Cells(i, 7).Value

End If

Next i

'You should also have conditional formatting that will highlight positive change in green and negative change in red.

'set row count
RowCount = Cells(Rows.Count, 12).End(xlUp).Row

'loop for yearly change conditional formatting

    For j = 2 To RowCount

        Yearly_Change = Cells(j, 12).Value

        If Yearly_Change >= 0 Then

        Cells(j, 12).Interior.ColorIndex = 4
        Else
        Cells(j, 12).Interior.ColorIndex = 3

        End If

    Next j



End Sub