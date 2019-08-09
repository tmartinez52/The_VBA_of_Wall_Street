Sub StockDataEasy()
    Dim Stock_Letter As String
    Dim Stock_Total As Double
    Stock_Total = 0

    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2

    For i = 2 To 760192

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

    Stock_Letter = Cells(i, 1).Value

    Stock_Total = Stock_Total + Cells(i, 7).Value

    Range("i" & Summary_Table_Row).Value = Stock_Letter

    Range("j" & Summary_Table_Row).Value = Stock_Total

    Summary_Table_Row = Summary_Table_Row + 1

    Stock_Total = 0

    Else
        Stock_Total = Stock_Total + Cells(i, 7).Value

    End If

    Next i
End Sub
