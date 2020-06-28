Attribute VB_Name = "Module1"
Sub ticker()

    ' ----------------------------------------
    ' LOOP THROUGH ALL SHEETS
    ' ----------------------------------------
    

    For Each ws In Worksheets
    
        ' ----------------------------------------
        ' SET VARIABLES
        ' ----------------------------------------
        

        ' Set value for holding the ticker symbol, opening price, closing price, yearly change, percent change, and total stock volume
        Dim ticker As String
        Dim opening_price As Double
        Dim closing_price As Double
        Dim yearly_change As Double
        Dim percent_change As Double
        Dim total_volume As Double
        Dim greatest_increase As Double
        Dim greatest_decrease As Double
        Dim greatest_total_volume As Double
        
     
        ' Set values for Ticker, Yearly Change, Percent Change, and Total Stock Volume columns
        ws.Range("I1").Value = "Ticker"
        ws.Range("J1").Value = "Yearly Change"
        ws.Range("K1").Value = "Percent Change"
        ws.Range("L1").Value = "Total Stock Volume"
        ws.Cells(2, 14).Value = "Greatest % Increase"
        ws.Cells(3, 14).Value = "Greatest % Decrease"
        ws.Cells(4, 14).Value = "Greatest Total Volume"
        ws.Cells(1, 15).Value = "Ticker"
        ws.Cells(1, 16).Value = "Value"
        
    
        ' Set initial values for variables
        opening_price = 0
        closing_price = 0
        yearly_change = 0
        percent_change = 0
        total_volume = 0
        
        ' Set values for colors
        color_green = 4
        color_red = 3
        
        
        ' Keep track of the location for each ticker in the summary table
        Dim Summary_Table_Row As Integer
        Summary_Table_Row = 2
    
        ' Set variable for last row
        lastrow = Cells(Rows.Count, 1).End(xlUp).Row
        ' MsgBox (lastrow)
    
        ' Loop through all rows
        For i = 2 To lastrow
        
        'Check if ticker is still the same, if not then set the value in the Ticker column
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
                
                 ' Add to the Total Stock Volume
                total_volume = total_volume + ws.Cells(i, 7).Value
                ' MsgBox ("Total Volume " & total_volume)
        
                'Set the ticker symbol value
                ticker = ws.Cells(i, 1).Value
                ' MsgBox ("Ticker" & ticker)
                
                ' Add to the Opening Price
                opening_price = opening_price + ws.Cells(i, 3).Value
                ' MsgBox ("Sum of opening price " & opening_price)
        
                ' Add to the Closing Price
                closing_price = closing_price + ws.Cells(i, 6).Value
                ' MsgBox ("Sum of closing price " & closing_price)
            
                            
                ' Calculation for obtaining Yearly Change
                yearly_change = closing_price - opening_price
                ' MsgBox ("yearly change " & yearly_change)
            
                ' Calculation for obtaining Percent Change
                percent_change = yearly_change / opening_price
                
                'MsgBox (percent_change)
            
                ' Print the ticker symbol in the Ticker column
                ws.Range("I" & Summary_Table_Row).Value = ticker
        
                ' Print the Yearly Change
                ws.Range("J" & Summary_Table_Row).Value = yearly_change
            
                    ' If the yearly change is positive, setting cell to green
                    If yearly_change > 0 Then
                    
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = color_green
                
                    Else
                    
                        ' Setting cell to red if yearly change is not positive
                        ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = color_red
                    
                    End If
            
                ' Print the Percent Change
                ws.Range("K" & Summary_Table_Row).Value = percent_change
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
            
                ' Print the Total Volume
                ws.Range("L" & Summary_Table_Row).Value = total_volume
        
                ' Add on to the summary table row
                Summary_Table_Row = Summary_Table_Row + 1
        
                ' Reset the totals
                opening_price = 0
                closing_price = 0
                yearly_change = 0
                percent_change = 0
                total_volume = 0
        
            Else
        
                ' Add to the opening price, closing price, and volume totals
                opening_price = opening_price + Cells(i, 3).Value
                closing_price = closing_price + Cells(i, 6).Value
                total_volume = total_volume + Cells(i, 7).Value

        
            End If
        
                
        Next i
        
       
        ' -----------------------------
        ' CHALLENGE PORTION
        ' -----------------------------
        
        ' Populate the greatest increase, descrease, and total volume
        
        ' Define new last row for new column data
        newlastrow = Cells(Rows.Count, 9).End(xlUp).Row
        
        ' Loop through each row
        For i = 2 To newlastrow
        
            ' Set greatest_increase to be the max of all values in Column J
            greatest_increase = Application.WorksheetFunction.Max(ws.Columns("J"))
            
            'Set greatest_decrease to be the min of all values in Column J
            greatest_decrease = Application.WorksheetFunction.Min(ws.Columns("J"))
            
            'Set greatest_total_volume to be the max of all values in Column L
            greatest_total_volume = Application.WorksheetFunction.Max(ws.Columns("L"))
            
            ' Compare to see if the current value in column J for the row equals the max (greatest_increase)
            If ws.Cells(i + 1, 10) = greatest_increase Then
            
                ' grab the value of the ticker symbol
                ticker = ws.Cells(i + 1, 9).Value
                
                ' set the values of the ticker and greatest increase
                ws.Range("O2").Value = ticker
                ws.Range("P2").Value = greatest_increase / 100
                ws.Range("P2").NumberFormat = "0.00%"
            
            End If
            
            
            ' Compare to see if the current value in column J for the row equals the min (greatest_decrease)
            If ws.Cells(i + 1, 10) = greatest_decrease Then
            
                ' grab the value of the ticker symbol
                ticker = ws.Cells(i + 1, 9).Value
                
                ' set the values of the ticker and greatest decreaase
                ws.Range("O3").Value = ticker
                ws.Range("P3").Value = greatest_decrease / 100
                ws.Range("P3").NumberFormat = "0.00%"
            
            End If
            
            ' Compare to see if the current value in column J for the row equals the greatest total volume
            If ws.Cells(i + 1, 12) = greatest_total_volume Then
            
                ' grab the value of the ticker symbol
                ticker = ws.Cells(i + 1, 9).Value
                
                ' set the values of the ticker and greatest total volume
                ws.Range("O4").Value = ticker
                ws.Range("P4").Value = greatest_total_volume
            
            
            End If
                
            
        Next i


    Next ws
    
End Sub


