Attribute VB_Name = "Module1"
Sub StockMarketTestData():
'The ticker symbol.
    LastRow = Cells(Rows.Count, "A").End(xlUp).Row
        Cells(1, 10).Value = "Ticker"
            For i = 2 To LastRow
                If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
                    Cells(i, 10).Value = TickerValue
                    Range("J" & 2 + j).Value = Cells(i, 1).Value
                    j = j + 1
                End If
            Next i
'Yearly change from opening price at the beginning of a given year to the closing price at the end of that year.
        Range("K1").Value = "Yearly Change"
            Dim YearlyChange As Double
            For i = 2 To LastRow
                If Cells(i + 1, 10).Value <> Cells(i, 10).Value Then
                    YearlyChange = Cells(i, 6).Value - Cells(i, 3).Value
                    Cells(i, 11).Value = YearlyChange
                End If
            Next i
'Conditional Formatting
            For i = 2 To LastRow
            If Cells(i, 11).Value < 0 Then
                Cells(i, 11).Interior.ColorIndex = 3
                ElseIf Cells(i, 11).Value > 0 Then
                Cells(i, 11).Interior.ColorIndex = 4
                Else: Cells(i, 11).Value = 0
                Cells(i, 11).Interior.ColorIndex = 6
                End If
            Next i
'The percent change from opening price at the beginning of a given year to the closing price at the end of that year.
        Range("L1") = "Percent Change"
            For i = 2 To LastRow
            Dim percentChange As Integer
                If Cells(i + 1, 10).Value <> Cells(i, 10).Value Then
                    YearlyChange = Cells(i, 6).Value - Cells(i, 3).Value
                    percentChange = Round((YearlyChange / (Cells(i, 3)) * 100), 2)
                    Cells(i, 12).Value = percentChange
                    Cells(i, 12).Value = "%" & percentChange
                End If
            Next i
'The total stock volume of the stock.
                    Cells(1, 13).Value = "Total Stock Volume"
                Dim TotalStock As Long
                    For i = 2 To LastRow
                        TotalStock = Cells(i, 6).Value * Cells(i, 7).Value
                        Cells(i, 13) = TotalStock
                    Next i
End Sub

