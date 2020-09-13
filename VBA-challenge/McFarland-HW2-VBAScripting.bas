Attribute VB_Name = "Module1"
Sub iterate():

    Dim i As Long
    
    Dim RowCount As Long
    RowCount = ActiveSheet.UsedRange.Rows.Count
    
    Dim RowNum As Long
    RowNum = 2
    
    Dim yearopen As Double
    Dim yearclose As Double
    Dim volume As Double
    
    Range("I:L").Clear
    
    
    For i = 2 To RowCount

        
        'checks to see if this is the first value for the ticker
        If Not Cells(i - 1, 1).Value = Cells(i, 1).Value Then
            yearopen = Cells(i, 3).Value
            volume = 0
        
        'checks to see if this is the last value for the ticker
        ElseIf Not Cells(i + 1, 1).Value = Cells(i, 1).Value Then
            Cells(RowNum, 9).Value = Cells(i, 1).Value
            yearclose = Cells(i, 6).Value
            Cells(RowNum, 10).Value = yearclose - yearopen
            
            'check to make sure no division by 0
            If Not yearopen = 0 Then
                Cells(RowNum, 11).Value = Format((yearclose - yearopen) / yearopen, "Percent")
                
            Else
                Cells(RowNum, 11).Value = Format(0, "Percent")
                
            End If
         
            volume = volume + Cells(i, 7).Value
            Cells(RowNum, 12).Value = volume
            volume = 0
            RowNum = RowNum + 1
            
        End If
        
        volume = volume + Cells(i, 7).Value

    Next i
End Sub
