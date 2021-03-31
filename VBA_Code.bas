Attribute VB_Name = "Module1"
Sub stock_count()

Dim Ticker As String
Dim Year_Open As Double
Year_Open = 0
Dim Year_Closed As Double
Year_Closed = 0
Dim Yearly As Double
Yearly = 0
Dim Percentage As Double
Percentage = 0
Dim Total_Stock_Vol As Double
Total_Stock_Vol = 0

    For Each ws In Worksheets
    
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Yearly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
    
        LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
        Table_Row = 2
        
        Year_Open = ws.Cells(2, 3).Value
        
        For i = 2 To LastRow
        
          If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
            
                Ticker = ws.Cells(i, 1).Value
                ws.Cells(Table_Row, 9).Value = Ticker
                
                Year_Closed = ws.Cells(i, 6).Value
                Yearly = Year_Closed - Year_Open
                ws.Cells(Table_Row, 10).Value = Yearly
                If Yearly > 0 Then
                    ws.Cells(Table_Row, 10).Interior.ColorIndex = 4
                Else
                    ws.Cells(Table_Row, 10).Interior.ColorIndex = 3
                End If
                
                If Year_Open <> 0 Then
                    Percentage = Yearly / Year_Open
                Else
                    ws.Cells(Table_Row, 11).Value = "Null"
                End If
                ws.Cells(Table_Row, 11).Value = Percentage
                ws.Columns(11).NumberFormat = "0.00%"
                
                Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
                ws.Cells(Table_Row, 12).Value = Total_Stock_Vol
                
                Year_Open = ws.Cells(i + 1, 3).Value
                Total_Stock_Vol = 0
                
                Table_Row = Table_Row + 1
                
            Else
                
                Total_Stock_Vol = Total_Stock_Vol + ws.Cells(i, 7).Value
                
             End If
                       
        Next i
        
    Next ws

End Sub

