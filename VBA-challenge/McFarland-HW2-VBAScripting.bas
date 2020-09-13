Attribute VB_Name = "Module1"
Sub iterate():

    Dim i As Long
    
    'Finds total number of rows used in sheet
    Dim RowCount As Long
    RowCount = ActiveSheet.UsedRange.Rows.Count
    
    'Defines output rows to increment
    Dim RowNum As Long
    RowNum = 2
    
    Dim yearopen As Double
    Dim yearclose As Double
    Dim volume As Double
    
    'Resets output range and applies headers
    Range("I:P").Clear
    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Price Change"
    Cells(1, 11).Value = "Percent Change"
    Cells(1, 12).Value = "Total Volume"
    Cells(1, 15).Value = "Ticker"
    Cells(1, 16).Value = "Value"
    Cells(2, 14).Value = "Greatest % Increase"
    Cells(3, 14).Value = "Greatest % Decrease"
    Cells(4, 14).Value = "Greatest Total Volume"
    
    'Loops through tickers
    For i = 2 To RowCount

        'Finds first value of ticker, defines year opening price, and resets total volume
        If Not Cells(i - 1, 1).Value = Cells(i, 1).Value Then
            yearopen = Cells(i, 3).Value
            volume = 0
        
        'Finds last value of ticker
        ElseIf Not Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            Cells(RowNum, 9).Value = Cells(i, 1).Value
            yearclose = Cells(i, 6).Value
            Cells(RowNum, 10).Value = yearclose - yearopen
            
            'Formats positive (green) and negative (red) price change
            If Cells(RowNum, 10).Value > 0 Then
                Cells(RowNum, 10).Interior.ColorIndex = 4
            
            ElseIf Cells(RowNum, 10).Value < 0 Then
                Cells(RowNum, 10).Interior.ColorIndex = 3
            
            End If
                
            'Displays percent change (no division by 0) and or "N/A" (division by 0)
            If Not yearopen = 0 Then
                Cells(RowNum, 11).Value = Format((yearclose - yearopen) / yearopen, "Percent")
                
            Else
                Cells(RowNum, 11).Value = "N/A"
                
            End If
            
            'Adds last total volume row per ticker, displays and resets
            volume = volume + Cells(i, 7).Value
            Cells(RowNum, 12).Value = volume
            volume = 0
            
            'Increments output row number
            RowNum = RowNum + 1
            
        End If
        
        'Sums volume until above elseif is triggered, where it is reset for next ticker
        volume = volume + Cells(i, 7).Value

        If Cells(i, 10).Value > Cells(i - 1, 10).Value Then
            Cells(i, 1).Value = Cells(2, 15).Value
            Cells(i, 10).Value = Cells(2, 16).Value
        
        End If
        

    Next i
    
    
    
    
End Sub
