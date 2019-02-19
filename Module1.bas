Attribute VB_Name = "Module1"
Sub Homework2()

For Each ws In Worksheets

    lastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
    Dim company_total As Double
    company_total = 0
    Dim company_counter As Integer
    company_counter = 2
    
    ws.Cells(1, 9) = "Ticker"
    ws.Cells(1, 10) = "Total Stock Volume"
    
    For i = 2 To lastRow
        
        If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
        ws.Cells(company_counter, 9) = ws.Cells(i, 1).Value
        company_total = 0
        company_counter = company_counter + 1
        
        Else
        company_total = company_total + ws.Cells(i + 1, 7).Value
        ws.Cells(company_counter, 10) = company_total
        End If
        
    Next i

Next ws

End Sub
