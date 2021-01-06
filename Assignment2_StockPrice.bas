Attribute VB_Name = "Module1"
Sub stock_volumn()

For Each ws In Worksheets
    Dim WorksheetName As String
    WorksheetName = ws.Name
    Sheets(ws.Name).Select


'clear columns
Columns("I:Q").Select
Selection.Clear

Columns("I:Q").EntireColumn.AutoFit
    Cells(1, 1).Select

'add headings
Cells(1, 9).Value = "Ticker"
Cells(1, 12).Value = "Total Stock Volumn"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"
Cells(2, 15).Value = "Greatest % Increase"
Cells(3, 15).Value = "Greatest % Decrease"
Cells(4, 15).Value = "Greatest Total Volumn"


'set an initial variable for hodling the stock tickers
Dim ticker As String

'set an initial variable for hodling the total stock per ticker
Dim volumn As Double
    volumn = 0

'set an initial variable for hodling the last row per ticker
Dim last_row As Long
last_row = Cells(Rows.Count, "A").End(xlUp).Row

'Keep track of the location for each ticker in the summary table
Dim summary_table_row As Integer
    summary_table_row = 2

Dim y As Double
Dim x As Double

y = 2
x = 2

'set an initial variable for holding the open price
Dim open_price As Double
    open_price = 0

'set an initial variable for hodling the close price
Dim close_price As Double
    close_price = 0

'set initial open price
open_price = Cells(y, 3).Value



'loop through all tickers
For i = 2 To last_row

'Check if we are still within the same stock ticker, if it is not...
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
    'set ticker
    ticker = Cells(i, 1).Value
    
    'add the ticker total
    volumn = volumn + Cells(i, 7).Value

    'print the ticker in the summary table
    Range("I" & summary_table_row).Value = ticker
    Range("L" & summary_table_row).Value = volumn
    
     ' Add one to the summary table row
    summary_table_row = summary_table_row + 1
    
    'reset the volumn
    volumn = 0
    
    Else
    volumn = volumn + Cells(i, 7).Value
     
    End If

Next i
    

'loop through all tickers
For i = 2 To last_row
    
    'check if the ticker is the same as the ticker on summary table, if it is......
    If Cells(i, 1).Value = Cells(x, 9).Value Then
    
    'set close price
    close_price = Cells(i, 6).Value
    
    'if it is not
    Else
    
     'if ticker is not the same as ticker on summary table then
     Cells(x, 10).Value = close_price - open_price
       
        'if close price is smaller and equal th\an zero, then
        If close_price <= 0 Or open_price = 0 Then
           
           'percent change equale to zero
          Cells(x, 11).Value = 0
            
            'if close price is larger than zero
            Else
            
            'calculate the percent change
            Cells(x, 11).Value = (close_price - open_price) / open_price
       
        End If
        Cells(x, 11).Style = "Percent"
 
    'fill color in the cell based on yearly change
    If Cells(x, 10).Value >= 0 Then
        Cells(x, 10).Interior.ColorIndex = 4
        Else
        Cells(x, 10).Interior.ColorIndex = 3
    End If

    'set open price
    open_price = Cells(i, 3).Value
    
    x = x + 1
   
    
    End If

Next i

'bonus part

Dim ticker_increase As String
Dim greatest_percent_increase As Double
Dim ticker_decrease As String
Dim greatest_percent_decrease As Double
Dim ticker_volumn As String
Dim greatest_volumn As Double

Dim LastRow_2 As Long
LastRow_2 = Cells(Rows.Count, 9).End(xlUp).Row
        
For x = 2 To LastRow_2
        
        
If Cells(x, 11).Value > greatest_percent_increase Then
            
ticker_increase = Cells(x, 9).Value
greatest_percent_increase = Cells(x, 11).Value
        
End If
        
        
If Cells(x, 11).Value < greatest_percent_decrease Then
            
ticker_decrease = Cells(x, 9).Value
greatest_percent_decrease = Cells(x, 11).Value
        
End If

If Cells(x, 12).Value > greatest_volumn Then
            
ticker_volumn = Cells(x, 9).Value
greatest_volumn = Cells(x, 12).Value
        
End If
        
Next x
        
Cells(2, 16).Value = ticker_increase
Cells(2, 17).Value = greatest_percent_increase
Cells(2, 17).Style = "Percent"
Cells(3, 16).Value = ticker_decrease
Cells(3, 17).Value = greatest_percent_decrease
Cells(3, 17).Style = "Percent"
Cells(4, 16).Value = ticker_volume
Cells(4, 17).Value = greatest_volumn


Next ws
End Sub





