Attribute VB_Name = "Module1"
Sub Ticker_Analyis()

    'Declare ws as Worksheet
        Dim ws As Worksheet
        
    'Create a loop for each worksheet in workbook
    For Each ws In ThisWorkbook.Worksheets
    
    'Create headers for calculated fields
        ws.[I1] = "Ticker"
        ws.[J1] = "Yearly Change"
        ws.[K1] = "Percent Change"
        ws.[L1] = "Total Stock Volume"
        
        ws.Columns("I:L").AutoFit
        ws.Columns("K:K").NumberFormat = "0.00%"
    
        row_count = ws.Cells(Rows.Count, "A").End(xlUp).row
        j = 2
        total = 0
        firstOpen = 0
        
    
    'To total volume per ticker
        For i = 2 To row_count
        
            If firstOpen = 0 Then
                firstOpen = ws.Cells(i, "C")
            End If
        
           If ws.Cells(i, 1) <> ws.Cells(i + 1, 1) Then
                total = total + ws.Cells(i, 7)
                ws.Cells(j, "I") = ws.Cells(i, 1)
                ws.Cells(j, "L") = total
                
                yearlyCh = ws.Cells(i, "F") - firstOpen
                ws.Cells(j, "J") = yearlyCh
                
                ws.Cells(j, "K") = yearlyCh / firstOpen
                
                If yearlyCh > 0 Then
                    ws.Cells(j, "J").Interior.ColorIndex = 4
                Else
                    ws.Cells(j, "J").Interior.ColorIndex = 3
                End If
                
            
                
                 j = j + 1
                 total = 0
                 firstOpen = 0
            End If
            
            
        Next i
        
    Next ws

End Sub

