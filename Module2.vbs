Attribute VB_Name = "Module2"
Sub stockreportWhighlow()

Dim ticker As String 'stock name
Dim start As Double 'opening price
Dim finish As Double 'closing price
Dim Highest As Double 'Highest % change
Dim Lowest As Double 'lowest % change
Dim HighTicker As String 'highest % change ticker
Dim lowticker As String 'lowesst % change ticker
Dim Volticker As String 'highest volume ticker
Dim MostVol As Double 'highest volume
Dim volume As Double 'stock volume (Single variable type was giving an incorrect value for total volume, works with double)
Dim rowcount As Long 'counter for while loop
Dim tickrow As Long 'counter for unique stock names



'reset highest/lowest/volume/rowcount
Highest = 0
Lowest = 0
MostVol = 0
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
    
    'replace ticker and value for highest % change if current ticker is higher than current highest
    If (finish - start) / start > Highest Then
        Highest = ((finish - start) / start)
        HighTicker = ticker
    End If
        
    'replace ticker and value for Lowest % change if current ticker is Lower than current Lowest
    If (finish - start) / start < Lowest Then
        Lowest = (finish - start) / start
        lowticker = ticker
    End If
    
    'replace ticker and value for highest volume if current ticker has higher volume than current highest
    If volume > MostVol Then
        MostVol = volume
        Volticker = ticker
    End If
       
Wend 'end of outer loop

'print the table for highest % increase etc.
Cells(2, 15) = "Greatest % Increase"
Cells(3, 15) = "Greatest % Decrease"
Cells(4, 15) = "Greatest Total Volume"
Cells(1, 16) = "Ticker"
Cells(1, 17) = "Value"
Cells(2, 16) = HighTicker
Cells(3, 16) = lowticker
Cells(4, 16) = Volticker
Cells(2, 17) = Format(Highest, "Percent")
Cells(3, 17) = Format(Lowest, "Percent")
Cells(4, 17) = MostVol


End Sub
