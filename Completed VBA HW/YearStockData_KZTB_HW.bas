Attribute VB_Name = "Module1"
Sub CalculateStockChanges()

    Dim TRows As Long
    Dim i As Long
    Dim PrevI As Long
    Dim TCount As Long
    
    Range("I1") = "Ticker"
    Range("J1") = "Yearly Change"
    Range("K1") = "Percent Change"
    Range("L1") = "Total Stock Volume"
    
    TRows = Cells(Rows.Count, 1).End(xlUp).Row
    TCount = 2
    PrevI = 2
    
    For i = 2 To TRows
        If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then
            FillStockChangesTable Cells(i, 1).Value, TCount, PrevI, i
            TCount = TCount + 1
            PrevI = i + 1
        End If
    Next i
    
End Sub
Sub FillStockChangesTable()

    Dim Ticker As String
    Dim RowStockVolTable As Long
    Dim iRow As Long
    Dim EndRow As Long
    
    Dim StockVol As Double
    Dim yChange As Double
    Dim pChange As Double
    StockVol = 0
    
    Range("I" & RowStockVolTable).Value = Ticker
    yChange = Range("F" & EndRow).Value - Range("C" & iRow).Value
    Range("J" & RowStockVolTable).Value = yChange
    If yChange >= 0 Then
            Range("J" & RowStockVolTable).Interior.ColorIndex = 4
        Else
            Range("J" & RowStockVolTable).Interior.ColorIndex = 3
        End If

    If Range("C" & iRow).Value <> 0 Then
        pChange = yChange / Range("C" & iRow).Value
        Range("K" & RowStockVolTable).Value = Format(pChange, "Percent")
    Else
        Range("K" & RowStockVolTable).Value = Format(0, "Percent")
    End If
    
    For i = iRow To EndRow
        If Cells(i, 1).Value = Ticker Then
            StockVol = StockVol + Cells(i, 7).Value
        End If
    Next i
    Range("L" & RowStockVolTable).Value = StockVol
    
End Sub
Sub PerIncrease()
    Dim TRows As Long
    Dim i As Long
    Dim MaxValue As Double
    Dim MaxTicker As String
    
    MaxValue = Range("K2").Value
    For i = 2 To TRows
        If Range("K" & i).Value >= MaxValue Then
            MaxValue = Range("K" & i).Value
            MaxTicker = Range("I" & i).Value
        End If
    Next i
    
    Range("P2").Value = MaxTicker
    Range("Q2").Value = Format(MaxValue, "Percent")
End Sub
Sub PDecrease()
    Dim TRows As Long
    Dim i As Long
    Dim MinValue As Double
    Dim MInTicker As String
    
    MinValue = Range("K2").Value
    For i = 2 To TRows
        If Range("K" & i).Value <= MinValue Then
            MinValue = Range("K" & i).Value
            MInTicker = Range("I" & i).Value
        End If
    Next i
    
    Range("P3").Value = MInTicker
    Range("Q3").Value = Format(MinValue, "Percent")
End Sub
