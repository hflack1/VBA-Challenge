Attribute VB_Name = "Module1"
Sub stockreport1year()

Dim ticker As String 'stock name
Dim start As Double 'opening price
Dim finish As Double 'closing price
Dim volume As Double 'stock volume (Single variable type was giving an incorrect value for total volume, works with double)
Dim rowcount As Long 'counter for while loop
Dim tickrow As Long 'counter for unique stock names


rowcount = 2
tickrow = 2

'add headers
Cells(1, 9) = "Ticker"
Cells(1, 10) = "Yearly Change"
Cells(1, 11) = "Percent Change"
Cells(1, 12) = "Total Stock Volume"

X:
While Cells(rowcount, 1) <> "" 'loop until cells are empty
    
    ticker = Cells(rowcount, 1) 'set next ticker name
    volume = 0 'reset volume for each ticker
    
    If Cells(rowcount, 3) = 0 Then
        rowcount = rowcount + 1
        GoTo X 'goto the beginning of this loop when the starting value of the stock is zero
    End If
    
    start = Cells(rowcount, 3) 'get year opening value
    
    While Cells(rowcount, 1) = ticker 'inner loop until ticker changes
        volume = volume + Cells(rowcount, 7) 'add daily volume to running total
        finish = Cells(rowcount, 6) 'grab closing price
        rowcount = rowcount + 1
        
    Wend 'end of inner loop
        
    'print individual stock information in new list
    Cells(tickrow, 9) = ticker 'ticker symbol
    Cells(tickrow, 10) = finish - start 'yearly change
    Cells(tickrow, 11) = Format((finish - start) / start, "Percent") 'yearly percent change
    Cells(tickrow, 12) = volume 'yearly volume
    
    'color formatting for % change column red for less than zero, green for more than zero
    If Cells(tickrow, 10) < 0 Then
        Cells(tickrow, 10).Interior.ColorIndex = 3 'color index 3 is red
    ElseIf Cells(tickrow, 10) > 0 Then
        Cells(tickrow, 10).Interior.ColorIndex = 4 'color index 4 is green
    End If
    
    tickrow = tickrow + 1
    
       
Wend 'end of outer loop


End Sub
