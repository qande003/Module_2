Sub Stock_Price()

 Dim Ticker_Name As String
    
    'Set an initial variable for holding the total stock volume

    Dim Ticker_Total_Volume As Double
    Ticker_Total_Volume = 0
    
    Dim Summary_Table_Row As Double
    Summary_Table_Row = 2
    
    'Set variables for yearly change and percentage change
    
    Dim Yearly_Change As Double
    Dim Percentage_Change As Double
    
    'Name ticker open and close
    
    Dim Ticker_Open_Price As Double
    Dim Ticker_Close_Price As Double
    
    'Manually name summary table headers"
    
    Range("M1").Value = "Ticker"
    Range("N1").Value = "Yearly Change"
    Range("O1").Value = "Percent Change"
    Range("P1").Value = "Total Stock Volume"
    
    'Final table variable
    Dim Greatest_Percent_Increase_Ticker As String
    Dim Greatest_Percent_Increase As Double
    Greatest_Percent_Increase = -10000000
    Dim Greatest_Percent_Decrease_Ticker As String
    Dim Greatest_Percent_Decrease As Double
    Greatest_Percent_Decrease = 10000000
    Greatest_Total_Volume = -1
    Dim Greatest_Total_Volume_Ticker As String
    
    
    'Final table
    Range("V1").Value = "Ticker"
    Range("W1").Value = "Value"
    Cells(2, 21).Value = "Greatest % Increase"
    Cells(3, 21).Value = "Greatest % Decrease"
    Cells(4, 21).Value = "Greatest Total Volume"
    
    Dim i As Long
    
    i = 2
    
    Do While Cells(i, 1).Value <> ""
        
        'Check if in the first row
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
        
            'Find earliest open for each ticker symbol
            Ticker_Open_Price = Cells(i, 3).Value
            
        
        End If
    
        'Check if we are still within the same ticker name, if it is not...
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

            'Set the ticker name
            Ticker_Name = Cells(i, 1).Value
     
            'Add to the ticker total
            Ticker_Total_Volume = Ticker_Total_Volume + Cells(i, 7).Value
            
            
            'Find last close for each ticker symbol
            Ticker_Close_Price = Cells(i, 6).Value
            
            'Find yearly change
            Yearly_Change = Ticker_Close_Price - Ticker_Open_Price
            
            Percent_Change = (Ticker_Close_Price - Ticker_Open_Price) / Ticker_Open_Price
            
            'Highlight postive yearly change green and negative change red
            If Yearly_Change >= 0 Then
                
                Range("N" & Summary_Table_Row).Interior.Color = vbGreen
                
            Else
                Range("N" & Summary_Table_Row).Interior.Color = vbRed
                
            End If
      
            'Print the ticker name in the Summary Table
            Range("M" & Summary_Table_Row).Value = Ticker_Name
        
            'Print the ticker Amount to the Summary Table
            Range("P" & Summary_Table_Row).Value = Ticker_Total_Volume
            
            'Yearly change (yearly close-yearly open)
            Range("N" & Summary_Table_Row).Value = Yearly_Change
            
            'Percentage change calculation and format to percent
            Range("O" & Summary_Table_Row).Value = Percent_Change
            
            Range("O" & Summary_Table_Row).NumberFormat = "0.00%"
            
            If (Percent_Change > Greatest_Percent_Increase) Then
                Greatest_Percent_Increase = Percent_Change
                Greatest_Percent_Increase_Ticker = Ticker_Name
                
            
            End If
                
            If (Percent_Change < Greatest_Percent_Decrease) Then
                Greatest_Percent_Decrease = Percent_Change
                Greatest_Percent_Decrease_Ticker = Ticker_Name
                
            End If
             
        
            'Find greatest ticker volume
            If (Ticker_Total_Volume > Greatest_Total_Volume) Then
                Greatest_Total_Volume = Ticker_Total_Volume
                Greatest_Total_Volume_Ticker = Ticker_Name
            
            
            End If
             
            'Add one to the summary table row
            Summary_Table_Row = Summary_Table_Row + 1
            
            'Reset the ticker Total
            Ticker_Total_Volume = 0
    
        'If the cell immediately following a row is the same ticker...

        Else
          
            ' Add to the Brand Total
            Ticker_Total_Volume = Ticker_Total_Volume + Cells(i, 7).Value
                
        End If
      
        i = i + 1
        
    Loop
    
    'End of for loop
    
    Range("W2").Value = Greatest_Percent_Increase
    Range("W2").NumberFormat = "0.00%"
    
    Range("W3").Value = Greatest_Percent_Decrease
    Range("W3").NumberFormat = "0.00%"
    Range("V2").Value = Greatest_Percent_Increase_Ticker
    Range("V3").Value = Greatest_Percent_Decrease_Ticker
    Range("W4").Value = Greatest_Total_Volume
    Range("V4").Value = Greatest_Total_Volume_Ticker

End Sub
