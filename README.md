# VBA-Challenge

'Module 2 challenge

Sub ticker_counter()

    'establish ticker variable, total stock volume, summary ticker
    Dim ticker As String
    Dim total_volume, open_price, close_price, yearly_change, percent_change As Double
    Dim ticker_total As Integer
    
 
 
    'set starting points
    total_volume = 0
    ticker_total = 2
    open_price = Range("C2").Value
    
    
    'set header labels
    Range("I1").Value = "Ticker"
    Range("J1").Value = "Yearly Change"
    Range("K1").Value = "Percent Change"
    Range("L1").Value = "Total Stock Volume"

    'set headers for % increases
    Range("P1").Value = "Ticker"
    Range("Q1").Value = "Value"
    Range("O2").Value = "Greatest % Increase"
    Range("O3").Value = "Greatest % Decrease"
    Range("O4").Value = "Greatest Total Volume"
    
    'Total row count
    lastrow = Cells(Rows.Count, 1).End(xlUp).Row

    'Loop for summary table (ticker, yearly change, percent change, total volume)
    For i = 2 To lastrow
    
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
            ticker = Cells(i, 1).Value
        
            total_volume = total_volume + Cells(i, 7).Value
        
            'Yearly change calculation
                close_price = Cells(i, 6).Value
        
                yearly_change = (close_price - open_price)
            
            'percent change calculation
                percent_change = (yearly_change / open_price)
            
    'print volumes to summary table breakdown
            Range("I" & ticker_total).Value = ticker
        
             Range("J" & ticker_total).Value = yearly_change
            
             Range("K" & ticker_total).Value = percent_change
            'change percent_change format
             Range("K" & ticker_total).NumberFormat = "0.00%"
            
            Range("L" & ticker_total).Value = total_volume
          
    'reset volume, price, and summary table
    
        total_volume = 0
        
        ticker_total = ticker_total + 1
        
         open_price = Cells(i + 1, 3)
        
    Else
        
        total_volume = total_volume + Cells(i, 7).Value
        
    End If
    
Next i

'Postive and Negative change conditional loop
  lastrow_summary = Cells(Rows.Count, 9).End(xlUp).Row
  
  For i = 2 To lastrow_summary
    If Cells(i, 10).Value > 0 Then
        Cells(i, 10).Interior.ColorIndex = 4
    Else
        Cells(i, 10).Interior.ColorIndex = 3
    End If
 Next i
 
 '% Increase, % Decrease, Greatest Total Volume Calculation
For i = 2 To lastrow_summary

'Greatest % Increase Calculation and Formatting
    If Cells(i, 11).Value = WorksheetFunction.Max(Range("K2:K" & lastrow_summary)) Then
        Cells(2, 16).Value = Cells(i, 9).Value
        Cells(2, 17).Value = Cells(i, 11).Value
        Cells(2, 17).NumberFormat = "0.00%"

'Greatest % Decrease Calculation and Formatting
    ElseIf Cells(i, 11).Value = WorksheetFunction.Min(Range("K2:K" & lastrow_summary)) Then
        Cells(3, 16).Value = Cells(i, 9).Value
        Cells(3, 17).Value = Cells(i, 11).Value
        Cells(3, 17).NumberFormat = "0.00%"

'Greatest Total Volume Calculation
    ElseIf Cells(i, 12).Value = WorksheetFunction.Max(Range("L2:L" & lastrow_summary)) Then
        Cells(4, 16).Value = Cells(i, 9).Value
        Cells(4, 17).Value = Cells(i, 12).Value
        
    End If
Next i

    

End Sub
