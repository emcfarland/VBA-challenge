Attribute VB_Name = "Module1"
Sub Summarize_Stock_Info():

    Dim i As Long
    Dim j As Long
        
    'Finds total number of rows used in sheet
    Dim InputRowCount As Long
    InputRowCount = ActiveSheet.UsedRange.Rows.Count
    
    'Defines output rows to increment
    Dim OutputRowCount As Long
    OutputRowCount = 2
    
    Dim OpeningPrice As Double
    Dim ClosingPrice As Double
    Dim TotalVolume As Double
    
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

    
    'Loops through input tickers
    For i = 2 To InputRowCount

        'Finds first value of ticker, defines year opening price, and resets total TotalVolume
        If Not Cells(i - 1, 1).Value = Cells(i, 1).Value Then
            OpeningPrice = Cells(i, 3).Value
            TotalVolume = 0
        
        'Finds last value of ticker
        ElseIf Not Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            Cells(OutputRowCount, 9).Value = Cells(i, 1).Value
            ClosingPrice = Cells(i, 6).Value
            Cells(OutputRowCount, 10).Value = ClosingPrice - OpeningPrice
            
            'Formats positive (green) and negative (red) price change
            If Cells(OutputRowCount, 10).Value > 0 Then
                Cells(OutputRowCount, 10).Interior.ColorIndex = 4
            
            ElseIf Cells(OutputRowCount, 10).Value < 0 Then
                Cells(OutputRowCount, 10).Interior.ColorIndex = 3
            
            End If
                
            'Displays percent change (no division by 0) and or "N/A" (division by 0)
            If Not OpeningPrice = 0 Then
                Cells(OutputRowCount, 11).Value = Format((ClosingPrice - OpeningPrice) / OpeningPrice, "Percent")
                
            Else
                Cells(OutputRowCount, 11).Value = "N/A"
                
            End If
            
            'Adds last total TotalVolume row per ticker, displays and resets
            TotalVolume = TotalVolume + Cells(i, 7).Value
            Cells(OutputRowCount, 12).Value = TotalVolume
            TotalVolume = 0
            
            'Increments output row number
            OutputRowCount = OutputRowCount + 1
            
        End If
        
        'Sums TotalVolume until above if or elseif is triggered, where it is reset for next ticker
        TotalVolume = TotalVolume + Cells(i, 7).Value
                
    Next i
        
    'Loops through output tickers
    For j = 2 To OutputRowCount
        
        'Outputs maximum % change, minimum % change, maximum total TotalVolume, and associated tickers
        If Cells(j, 11).Value = Application.WorksheetFunction.Max(Range("K:K")) Then
            Cells(2, 15).Value = Cells(j, 9).Value
            Cells(2, 16).Value = Format(Cells(j, 11).Value, "Percent")
            
        ElseIf Cells(j, 11).Value = Application.WorksheetFunction.Min(Range("K:K")) Then
            Cells(3, 15).Value = Cells(j, 9).Value
            Cells(3, 16).Value = Format(Cells(j, 11).Value, "Percent")
            
        ElseIf Cells(j, 12).Value = Application.WorksheetFunction.Max(Range("L:L")) Then
            Cells(4, 15).Value = Cells(j, 9).Value
            Cells(4, 16).Value = Cells(j, 12).Value
            
        End If
    
    Next j
    Application.ActiveSheet.Columns("I:P").AutoFit

End Sub
