Attribute VB_Name = "Module2"
Sub multiple_year_stock2():

    'Declare and set worksheet
    Dim ws As Worksheet
    
    'Set new variables
    Dim open_price As Double
    open_price = 0
    Dim close_price As Double
    close_price = 0
    Dim price_change As Double
    price_change = 0
    Dim price_change_percent As Double
    price_change_percent = 0
    Dim Yearly_Change As Double
    Yearly_Change = 0
    Dim Percent_Change As Double
    Percent_Change = 0
    Dim Total_Stock_Volume As LongLong
    Total_Stock_Volume = 0
    
    
    
    'Set initial and last row for worksheet
    Dim Lastrow As Long
    Dim i As Long
    Dim j As Long
    
    'Define Ticker variable
    Dim Ticker As String
    
    'Dim Ticker_volume As Double
    
    'Do loop of current worksheet to Lastrow
    Dim TickerRow As Long
    
    'Loop through all stocks for one year
    For Each ws In Worksheets
    
        'Create the column headings
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change in $"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        
        ws.Range("P1").Value = "Ticker"
        ws.Range("Q1").Value = "Value"
        ws.Range("O2").Value = "Greatest % Increase"
        ws.Range("O3").Value = "Greatest % Decrease"
        ws.Range("O4").Value = "Greatest Total Volume"
        
        'Define Ticker variable
        Ticker = " "
        Total_Stock_Volume = 0
        TickerRow = 2

        'Create variable to hold stock volume
        'Dim stock_volume As Double
        'stock_volume = 0
        
        'Define Lastrow of worksheet
        Lastrow = ws.Cells(Rows.Count, "A").End(xlUp).Row
        
        'MsgBox (Lastrow)
        
        'Do loop of current worksheet to Lastrow
        'Start = 2
        
             
       open_price = ws.Cells(2, 3).Value
       
        
        For i = 2 To Lastrow
     

            'Ticker symbol output
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                'MsgBox (TickerRow)
            
                Ticker = ws.Cells(i, 1).Value

                'Yearly_Change
                Yearly_Change = ws.Cells(i, 6).Value - open_price
                
                
                    'Fill "Yearly Change", i.e. Yearly_Change with Green and Red colors
                    If (Yearly_Change >= 0) Then
                     'Fill column with GREEN color - good
                     ws.Cells(TickerRow, "J").Interior.ColorIndex = 4
                    Else
                    'Fill column with RED color - bad
                     ws.Cells(TickerRow, "J").Interior.ColorIndex = 3
                    End If
                            
                
                'MsgBox (Yearly_Change)


                ' Calculate Percent Change
                Percent_Change = (Yearly_Change / open_price)
                    ws.Cells(TickerRow, "K").NumberFormat = "0.00%"

                'display
                ws.Cells(TickerRow, "I").Value = Ticker
                ws.Cells(TickerRow, "J").Value = Yearly_Change
                ws.Cells(TickerRow, "K").Value = Percent_Change
                ws.Cells(TickerRow, "L").Value = Total_Stock_Volume

                open_price = ws.Cells(i + 1, 3).Value
                Total_Stock_Volume = 0
                TickerRow = TickerRow + 1
                
                                                       

            Else
                Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
            
            End If
                 

        Next i
    
                 Dim Greatest_Increase As Double
                 Greatest_Percentage_Increase = 0
                 Dim Greatest_Decrease As Double
                 Greatest_Percentage_Decrease = 0
                 'Dim Greatest_Total_Volume As Long
                 'Greatest_Total_Volume = 0
                 'Dim LastRoww As Long
                 'Dim PercentChangeRange As Double
                 'PercentChangeRange = 0
                 'Dim TickerRange As Double
                 
                 
                
                'LastRoww = ws.Cells(Rows.Count, "I").End(xlUp).Row
                PercentChangeRange = ws.Range("K2:K3001")
                TickerRange = ws.Range("I2:I3001")
                GreatestVolume = ws.Range("L2:L3001")
                
                    
                'Greatest Increase %
                 Greatest_Percentage_Increase = Application.WorksheetFunction.Max(PercentChangeRange)
                 ws.Cells(2, 17).Value = Greatest_Percentage_Increase
                 ws.Cells(2, 17).NumberFormat = "0.00%"
                 
                'Greatest Decrease %
                 Greatest_Percentage_Decrease = Application.WorksheetFunction.Min(PercentChangeRange)
                 ws.Cells(3, 17).Value = Greatest_Percentage_Decrease
                 ws.Cells(3, 17).NumberFormat = "0.00%"
                 
                 'Greatest Total Volume
                 Greatest_Total_Volume = Application.WorksheetFunction.Max(GreatestVolume)
                 ws.Cells(4, 17).Value = Greatest_Total_Volume
                 ws.Cells(4, 17).NumberFormat = "General"
         
                 For j = 2 To 3001
                 
                 Dim maxTicker As String
                 Dim minTicker As String
                 Dim Total_Vol_Ticker As String
                                 
                 If ws.Cells(j, 11).Value = Greatest_Percentage_Increase Then
                    maxTicker = ws.Cells(j, 9).Value
                    ws.Cells(2, 16).Value = maxTicker
                    
                 ElseIf ws.Cells(j, 11).Value = Greatest_Percentage_Decrease Then
                    minTicker = ws.Cells(j, 9).Value
                    ws.Cells(3, 16).Value = minTicker
                    
                 ElseIf ws.Cells(j, 12).Value = Greatest_Total_Volume Then
                    Total_Vol_Ticker = ws.Cells(j, 9).Value
                    ws.Cells(4, 16).Value = Total_Vol_Ticker
                 End If
                 
             Next j

    Next ws


End Sub

