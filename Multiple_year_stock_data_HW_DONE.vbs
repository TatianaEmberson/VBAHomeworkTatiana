Sub Alphabetical_testing()

'identify variables'

Dim ticker As String
Dim opening_price As Double
Dim closing_price As Double
Dim yearly_change As Double
Dim percentage_change As Double
Dim total_stock_volume As Double

'Loop for all worksheets'

For Each ws In Worksheets

lastRowState = ws.Cells(Rows.Count, "A").End(xlUp).Row

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Stock Volume"

' Initialize variables for each column'
 
    number_tickers = 0
    ticker = ""
    yearly_change = 0
    opening_price = 0
    percent_change = 0
    total_stock_volume = 0
    
'Loop tickers'

    For i = 2 To lastRowState
    
    
        ticker = Cells(i, 1).Value
        
'Define opening price for the ticker'
        
        If opening_price = 0 Then
            opening_price = Cells(i, 3).Value
        End If
        
' Add up the total stock volume values for a ticker.'

        total_stock_volume = total_stock_volume + Cells(i, 7).Value
    

    If Cells(i + 1, 1).Value <> ticker Then
 
     number_tickers = number_tickers + 1
     Cells(number_tickers + 1, 9) = ticker
            
'Yearly change value'

closing_price = Cells(i, 6)
    
yearly_change = closing_price - opening_price

' Add yearly change'
            
    Cells(number_tickers + 1, 10).Value = yearly_change
    
' Shading yearly changes'
         
         If yearly_change > 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 4
            ElseIf yearly_change < 0 Then
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 3
            Else
                Cells(number_tickers + 1, 10).Interior.ColorIndex = 6
            End If
  
' Calculate percent change value for ticker'

            If opening_price = 0 Then
                percent_change = 0
            Else
                percent_change = (yearly_change / opening_price)
            End If
  ' Format the percent_change value as a percent'
  
            Cells(number_tickers + 1, 11).Value = Format(percent_change, "Percent")
            
   ' Set opening price back to 0'
   
            opening_price = 0
            
    ' Add total stock volume value'
    
            Cells(number_tickers + 1, 12).Value = total_stock_volume
            
    ' Set total stock volume back to 0'
    
            total_stock_volume = 0
         
         End If
Next i
Next ws


            
End Sub
