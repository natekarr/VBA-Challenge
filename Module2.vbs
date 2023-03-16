Attribute VB_Name = "Module1"
Sub Stockorganizer():
      lastRow = Cells(Rows.Count, 1).End(xlUp).row
      Dim i, j, row, WS_Count As Integer
      Dim totalVol, greatestIncrease, greatestDecrease, greatestVolume As LongLong
      Dim yearChange, percentChange, openPrice, closePrice As Double
      Dim Ticker, NextTicker As String
      Dim Current As Worksheet
      WS_Count = ActiveWorkbook.Worksheets.Count
  For j = 1 To WS_Count
      Worksheets(j).Activate
      'Setting up headers and column spacing
      Range("I1").Value = "Ticker"
      Range("J1").Value = "Yearly Change"
      Columns("J").ColumnWidth = 13
      Range("K1").Value = "Percent Change"
      Columns("K").ColumnWidth = 14
      Range("L1").Value = "Total Stock Volume"
      Columns("L").ColumnWidth = 18
      Columns("O").ColumnWidth = 21
      Range("O2").Value = "Greatest % Increase"
      Range("O3").Value = "Greatest % Decrease"
      Range("O4").Value = "Greatest Total Volume"
      Range("P1").Value = "Ticker"
      Range("Q1").Value = "Value"
    
      row = 2
      totalVol = 0
      openPrice = Cells(2, 3).Value
      greatestIncrease = 0
      greatestDecrease = 0
      greatestVolume = 0
    
      For i = 2 To lastRow
        Ticker = Cells(i, 1).Value
        totalVol = totalVol + Cells(i, 7).Value
        NextTicker = Cells(i + 1, 1).Value
        If Ticker <> NextTicker Then
          closePrice = Cells(i, 6).Value
          yearChange = closePrice - openPrice
          Cells(row, 9).Value = Ticker
          Cells(row, 10).Value = yearChange
          If yearChange < 0 Then
            Cells(row, 10).Interior.ColorIndex = 3
          Else
            Cells(row, 10).Interior.ColorIndex = 4
          End If
          percentChange = FormatPercent(yearChange / openPrice)
          Cells(row, 11).Value = percentChange
          Cells(row, 12).Value = totalVol
          If totalVol > greatestVolume Then
            greatestVolume = totalVol
          End If
          If (yearChange / openPrice) < 0 Then
            If (yearChange / openPrice) < greatestDecrease Then
              greatestDecrease = yearChange / openPrice
            End If
          Else
            If (yearChange / openPrice) > greatestIncrease Then
              greatestIncrease = yearChange / openPrice
            End If
          End If
          openPrice = Cells(i + 1, 3).Value
          row = row + 1
          totalVol = 0
        End If
      Next i
    
      Range("Q2").Value = FormatPercent(greatestIncrease)
      Range("Q3").Value = FormatPercent(greatestDecrease)
      Range("Q4").Value = greatestVolume
    Next j
End Sub
