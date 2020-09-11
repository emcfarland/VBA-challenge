Attribute VB_Name = "Module1"
Sub iterate():

    Dim i As Long
    RowCount = ActiveSheet.UsedRange.Rows.Count
    Dim ticker As String
    Dim RowNum As Long
    RowNum = 2
    Dim PrevRow As Long
    PrevRow = RowNum - 1
    Dim summary As String
    
        
    For i = 2 To RowCount
        ticker = Cells(i, 1).Value
        summary = Cells(RowNum, 9).Value
        If Not ticker = Cells(PrevRow, 9).Value Then
            summary = ticker
            RowNum = RowNum + 1
        Else
        
        End If
    Next


End Sub
