Attribute VB_Name = "Module1"
Sub stocks()
   
     'Print Summary table headers
     Cells(1, 9).Value = "Summary Ticker"
     Cells(1, 10).Value = "Total Volume"
     Cells(1, 11).Value = "Yearly Change"
     Cells(1, 12).Value = "Percent Change"
     
     ' Set an initial variable for holding the ticker
     Dim ticker As String
     
     ' Set an initial variable for holding the total volume per stock
     Dim total_volume As Double
     total_volume = 0
     
     ' Set the initial variable for holding the price change
     Dim price_change As Double
     price_change = 0
          
     'Set the initial variable for holding the Summary_Table_Row
     Dim Summary_Table_Row As Integer
     Summary_Table_Row = 2
       
     'Set the initial variable for holding the column variable
     Dim Column As Integer
     Column = 1
     
     'Set opened price value in the spreadsheet
      price_opened = Cells(2, Column + 2).Value
      
     ' Determine the Last Row
     Dim LastRow As Long
     LastRow = Cells(Rows.Count, 2).End(xlUp).Row
                            
       'Loop through all stocks
        For i = 2 To LastRow
            
            'Keep calculating the total volume of each stock while looping through
            'the stocks with the same ticker
            total_volume = total_volume + Cells(i, 7).Value
            
            'Compare each following ticker with the previous one and and
            'if they differ perform the actions below
            If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
                ' Set the ticker value
                ticker = Cells(i, 1).Value
                
                ' Print the tickers in the Summary Table
                Range("I" & Summary_Table_Row).Value = ticker
                
                ' Print the total volume in the Summary Table
                Range("J" & Summary_Table_Row).Value = total_volume
                
                'Set the price closed
                price_closed = Cells(i, Column + 5).Value
                
                'Calculate yearly price change
                price_change = price_closed - price_opened
                
                'In order to avoid  the "cannot divide by zero" error use conditional
                If price_opened <> 0 Then
                    'Calculate the percent changed, assign it to the value
                    percent_changed = price_change / price_opened
                    
                    ' Print the price difference in the Summary Table
                    Range("K" & Summary_Table_Row).Value = price_change
                 End If
                
                'Prin and format the pecent_change value
                Range("L" & Summary_Table_Row).Value = percent_changed
                Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
                        
                ' Add one to the summary table row
                 Summary_Table_Row = Summary_Table_Row + 1
                 
                ' Reset the Brand Total
                total_volume = 0
                
                'Set initial opened price
                price_opened = Cells(i + 1, Column + 2).Value
    
           End If
           
        Next i
        
        'Set the red and green colors for negative and positive values
        ' in the yearly change row appropriately
        For j = 2 To LastRow
        
            If Cells(j, Column + 10).Value > 0 Then
                Cells(j, Column + 10).Interior.ColorIndex = 4
            Else
                Cells(j, Column + 10).Interior.ColorIndex = 3
            End If
         Next j
         
         
    'Set the headers for the third table
    Cells(2, 15).Value = "Greatest % increase"
    Cells(3, 15).Value = "Greatest % Decrease"
    Cells(4, 15).Value = "Greatest total volume"
    Cells(1, 16).Value = "Ticker"
    Cells(1, 17).Value = "Value"
    
    'Update the value for the last row
    LastRow = Cells(Rows.Count, 12).End(xlUp).Row
    
    'Define the max values for volume and procent change and min value for percent chagne
    Dim total_volume_maxValue As Variant
    total_volume_maxValue = Application.WorksheetFunction.Max(Range("J2:J" & LastRow))
    
    Dim percent_change_maxValue As Variant
    percent_change_maxValue = Application.WorksheetFunction.Max(Range("L2:L" & LastRow))
    
    Dim percent_change_minValue As Variant
    percent_change_minValue = Application.WorksheetFunction.Min(Range("L2:L" & LastRow))
            
    'Loop through the defined rows to find the max values for volume and percent change
    'and min value for percent change, format and print values to the table
    For k = 2 To LastRow
    
        If Cells(k, 12).Value = percent_change_maxValue Then
            Range("Q2").Value = percent_change_maxValue
            Range("Q2").NumberFormat = "0.00%"
            Range("P2").Value = Cells(k, 9).Value
            
       ElseIf Cells(k, 12).Value = percent_change_minValue Then
            Range("Q3").Value = percent_change_minValue
            Range("Q3").NumberFormat = "0.00%"
            Range("P3").Value = Cells(k, 9).Value
        
        ElseIf Cells(k, 10).Value = total_volume_maxValue Then
           Range("Q4").Value = total_volume_maxValue
           Range("P4").Value = Cells(k, 9).Value
        End If
        
    Next k
                     
End Sub
