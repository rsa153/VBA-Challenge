Attribute VB_Name = "Module1"
Sub Stocks()

    ' loop through each worksheet
    For Each ws In Worksheets
    
        ' set summary table columns in every worksheet
        ws.Range("I1") = "Ticker"
        ws.Range("J1") = "Yearly Change"
        ws.Range("K1") = "Percent Change"
        ws.Range("L1") = "Total Stock Volume"
    
        ' Set variables
        Dim Ticker As String
        Dim YearlyChange As Double
        Dim PercentChange As Double
        Dim Volume As Double
    
        Dim StockOpen As Double
        Dim StockClose As Double
    
        Dim Summary_Table_Row As Double
        Summary_Table_Row = 2
    
        ' find how many rows contain data in a worksheet that contains data in column A in each worksheet
        Dim lastrow As Double
        lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    
        ' Set variable Volume to 0
        Volume = 0

        ' Loop through all row starting at row 2 and ending at the last row
        For i = 2 To lastrow

             If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1) Then
                Ticker = ws.Cells(i, 1).Value
                Volume = Volume + ws.Cells(i, 7).Value

                ws.Range("I" & Summary_Table_Row).Value = Ticker
                ws.Range("L" & Summary_Table_Row).Value = Volume

                StockClose = ws.Cells(i, 6)

                If StockOpen = 0 Then
                    YearlyChange = 0
                    PercentChange = 0
                Else:
                    YearlyChange = StockClose - StockOpen
                    PercentChange = (StockClose - StockOpen) / StockOpen
                End If
    
                ws.Range("J" & Summary_Table_Row).Value = YearlyChange
                ws.Range("K" & Summary_Table_Row).Value = PercentChange
                ws.Range("K" & Summary_Table_Row).Style = "Percent"
                ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"

                Summary_Table_Row = Summary_Table_Row + 1

            ElseIf ws.Cells(i - 1, 1).Value <> ws.Cells(i, 1) Then
                 StockOpen = ws.Cells(i, 3)

            Else: Volume = Volume + ws.Cells(i, 7).Value

            End If

    Next i

    ' conditional formatting
    For i = 2 To lastrow

        If ws.Range("J" & i).Value > 0 Then
            ws.Range("J" & i).Interior.ColorIndex = 4

        ElseIf ws.Range("J" & i).Value < 0 Then
        ws.Range("J" & i).Interior.ColorIndex = 3
        
        End If

    Next i
    
Next ws
End Sub
