
Sub Data_Stock_Hard():

 'Defining Variables for "worksheet", "sheet count" and "lastrow of each sheet".

 Dim Sht As Worksheet
 Dim Sht_Count As Integer
 Dim lastrow As Double


 'Defining variables for Consolidated rows for ticker symbol with their Total stock volume.

 Dim Ticker_Row As Integer
 Dim Stock_TotalVolume As Double


 'Defining variables for calculating Yearly Change(Stock open value - Stock Closing value) and its percent.

 Dim Year_OpenVal As Double
 Dim Year_CloseVal As Double
 Dim Yearly_Change As Double
 Dim Yearly_Change_Percent As Double

 'Defining Variables for "Greatest Increase %", "Greatest Decrease %" and "Greatest stock volume".

 Dim Gpercent_Increase As Double
 Dim Gpercent_Decrease As Double
 Dim GTotal_Volume As Double


 'Gets count of sheets for Current Workbook

 Sht_Count = ActiveWorkbook.Worksheets.Count


    'Loops through each Worksheet.
  
    For Each Sht In Worksheets

     
     'Setting Headers for each Sheet.
    
     Sht.Cells(1, 9).Value = "Ticker"
     Sht.Cells(1, 10).Value = "Yearly Change"
     Sht.Cells(1, 11).Value = "Percent Change"
     Sht.Cells(1, 12).Value = "Total Stock Volume"
    
     Sht.Cells(2, 15).Value = "Greatest % Increase"
     Sht.Cells(3, 15).Value = "Greatest % Decrease"
     Sht.Cells(4, 15).Value = "Greatest Total Volume"
    
     Sht.Cells(1, 16).Value = "Ticker"
     Sht.Cells(1, 17).Value = "Value"
     Year_OpenVal = Sht.Cells(2, 3)
    
     'Setting initial values for "Stock volume", "Ticker Symbol Row", "Greatest Increase/Decrease percent" and "Greatest stock volume".
    
     Stock_TotalVolume = 0
     Ticker_Row = 2
     Gpercent_Increase = 0
     Gpercent_Decrease = 0
     GTotal_Volume = 0
    
    'Gets the count(Variable) of last row for every sheet.
    
     lastrow = Cells(Rows.Count, 1).End(xlUp).Row
      
     'Loops through each row of the sheet.

        For i = 2 To lastrow
         
        
          'Condition when Ticker Symbol matches with symbol in next row.
        
          If (Sht.Cells(i, 1).Value = Sht.Cells(i + 1, 1).Value) Then
              Stock_TotalVolume = Stock_TotalVolume + Sht.Cells(i, 7).Value

                'conditional to pick a non-zero stock open value
                If (Year_OpenVal = 0) Then
                    Year_OpenVal = Sht.Cells(i + 1, 3).Value
                End If
                
          'Condition when all data related to one Ticker is consolidated and found a new Ticker Symbol in next row.
          
          Else
              Stock_TotalVolume = Stock_TotalVolume + Sht.Cells(i, 7).Value
              Sht.Cells(Ticker_Row, 9).Value = Sht.Cells(i, 1).Value
        
        
              'Assigning stock closing value once the data related to that stock is found at end of the year.
            
              Year_CloseVal = Sht.Cells(i, 6).Value
            
        
              'Calculating difference between stock's open value at year starting and closing value at year ending.
            
               Yearly_Change = (Year_CloseVal - Year_OpenVal)
               Sht.Cells(Ticker_Row, 10).Value = Yearly_Change
               
               'Conditional formatting for yearly_change
               
               If (Sht.Cells(Ticker_Row, 10).Value > 0) Then
                   Sht.Cells(Ticker_Row, 10).Interior.ColorIndex = 4
                   
               ElseIf (Sht.Cells(Ticker_Row, 10).Value < 0) Then
                   Sht.Cells(Ticker_Row, 10).Interior.ColorIndex = 3
             
               ElseIf (Sht.Cells(Ticker_Row, 10).Value = 0) Then
                   Sht.Cells(Ticker_Row, 10).Interior.ColorIndex = 0
               
               End If
               
               'Accounting exception for stockvalues opening with "0".
            
               If (Year_OpenVal = 0) Then
                   Sht.Cells(Ticker_Row, 11).Value = 0
        
               Else
                   Yearly_Change_Percent = (Yearly_Change / Year_OpenVal) * 100
                   Sht.Cells(Ticker_Row, 11).Value = Yearly_Change_Percent
                   Sht.Cells(Ticker_Row, 11).Value = Round(Yearly_Change_Percent, 2)
                   Sht.Cells(Ticker_Row, 11).NumberFormat = "0.00\%"
               End If
               
               
               'Conditional formatting for yearly_percent
               
               If (Sht.Cells(Ticker_Row, 11).Value > 0) Then
                   Sht.Cells(Ticker_Row, 11).Interior.ColorIndex = 4
                   
               ElseIf (Sht.Cells(Ticker_Row, 10).Value < 0) Then
                   Sht.Cells(Ticker_Row, 11).Interior.ColorIndex = 3
             
               ElseIf (Sht.Cells(Ticker_Row, 10).Value = 0) Then
                   Sht.Cells(Ticker_Row, 11).Interior.ColorIndex = 0
               
               End If
            
            
                'Conditional to pick Greatest percent Increase.
            
                If (Sht.Cells(Ticker_Row, 11).Value > Gpercent_Increase) Then
                    Sht.Cells(2, 16).Value = Sht.Cells(Ticker_Row, 9).Value
                    Sht.Cells(2, 17).Value = Sht.Cells(Ticker_Row, 11).Value
                    Gpercent_Increase = Sht.Cells(Ticker_Row, 11).Value
                    Sht.Cells(2, 17).NumberFormat = "0.00\%"
            
                End If
            
                'Conditional to pick Greatest percent Decrease.
            
                If (Sht.Cells(Ticker_Row, 11).Value < Gpercent_Decrease) Then
                    Sht.Cells(3, 16).Value = Sht.Cells(Ticker_Row, 9).Value
                    Sht.Cells(3, 17).Value = Sht.Cells(Ticker_Row, 11).Value
                    Gpercent_Decrease = Sht.Cells(Ticker_Row, 11).Value
                    Sht.Cells(3, 17).NumberFormat = "0.00\%"
                
                End If
            
            
                'Writing Total stock volume for a stock in loop.
            
                Sht.Cells(Ticker_Row, 12).Value = Stock_TotalVolume
             
                'Conditional to pick Greatest stock total volume
                
                If (Sht.Cells(Ticker_Row, 12).Value > GTotal_Volume) Then
                    Sht.Cells(4, 16).Value = Sht.Cells(Ticker_Row, 9)
                    Sht.Cells(4, 17).Value = Sht.Cells(Ticker_Row, 12)
                    GTotal_Volume = Sht.Cells(Ticker_Row, 12).Value
            
                End If
            
                'Pointing to next desired row where my next iteration writes data.
            
                Ticker_Row = Ticker_Row + 1
            
                'Resetting stock total value to "0" and setting initial reading value for next iteration
            
                Stock_TotalVolume = 0
                Year_OpenVal = Sht.Cells(i + 1, 3)
            
            End If

        Next i

    Next

   
End Sub



