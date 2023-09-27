Sub run()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call stocks
    Next
    Application.ScreenUpdating = True
End Sub

Sub stocks():
        Dim symbol As String
        Dim yearly_change, init, percent_change, total_stock, opening_price, closing_price As Double
        Dim symbol_table As Integer
        
        Dim GI, GD, GTV, Num As Double
        Dim ticker, GDT, GIT, GTVT As String
        GD = Range("K2").Value
        GI = Range("K2").Value
        GTV = Range("L2").Value
        
        
        symbol_table = 2
        Cells(1, 9).Value = "Ticker"
        Cells(1, 10).Value = "Yaerly Change"
        Range("k1").Value = "Percent Change"
        Range("L1").Value = "Total Stock"
        
        Range("P2").Value = "Ticker"
        Range("Q2").Value = "Value"
        Range("O3").Value = "Greatest % Increase"
        Range("O4").Value = "Greatest % Decrease"
        Range("O5").Value = "Greatest Total Volume"
    
        
        
        n = Cells(Rows.Count, 1).End(xlUp).Row
        init = 2

        For I = 2 To n
        
        
        
        'check if we're still within the same symbol, if its not...
        
        If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
        'set new sybol
        symbol = Cells(I, 1).Value
        'add to the total_stock
        total_stock = total_stock + Cells(I, 7).Value
        
        'print the symbol in the summary Table
            Range("I" & symbol_table).Value = symbol
            
            
         'calculate the yearly change
            opening_price = Cells(init, 3).Value
            closing_price = Cells(I, 6).Value
            yearly_change = closing_price - opening_price
         
         'print the yearly change
         
            Range("J" & symbol_table).Value = yearly_change
            
         'set the color
        
            If yearly_change < 0 Then
                Cells(symbol_table, 10).Interior.ColorIndex = 3
           ElseIf yearly_change > 0 Then
               Cells(symbol_table, 10).Interior.ColorIndex = 4
            End If
            
         'calculate percentage change
            percent_change = (yearly_change / opening_price) * 100

         'print the PERCENTAGE change
            Range("K" & symbol_table).Value = percent_change
         'print the total_stock
            Range("L" & symbol_table).Value = total_stock
         'add one to the symbol_table
            symbol_table = symbol_table + 1
          'reset the total_stock
          
          total_stock = 0
          init = I + 1
        
        'if the following cell is the same symbol...
            Else
               '   add the symbol total
                   total_stock = total_stock + Cells(I, 7).Value
            End If
        Next I
        
        'resume chart
        m = Cells(Rows.Count, 11).End(xlUp).Row
        For I = 2 To m
            If GI < Cells(I, 11) Then
                GI = Cells(I, 11)
                Range("Q3").Value = GI & "%"
                GIT = Cells(I, 9).Value
                Range("P3").Value = GIT
            End If
            If GD > Cells(I, 11) Then
                GD = Cells(I, 11)
                Range("Q4").Value = GD & "%"
                GDT = Cells(I, 9).Value
                Range("P4").Value = GDT
                
            End If
        Next I
        
        'greatest total value
         o = Cells(Rows.Count, 12).End(xlUp).Row
         For I = 2 To o
            If GTV < Cells(I, 12) Then
                GTV = Cells(I, 12)
                Range("Q5").Value = GTV
                GTVT = Cells(I, 9).Value
                Range("P5").Value = GTVT
            End If
          Next I
        
End Sub