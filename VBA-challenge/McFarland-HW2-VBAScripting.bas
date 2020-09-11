Attribute VB_Name = "Module1"
Sub iterate():

    Dim i As Long
    RowCount = ActiveSheet.UsedRange.Rows.Count
    
    For i = 1 To RowCount
        Cells(i, 9).Value = i + 1
    Next

End Sub
