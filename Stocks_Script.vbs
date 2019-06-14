Sub Stocks()
  
Dim ws As Worksheet
Dim starting_ws As Worksheet
Set starting_ws = ActiveSheet

For Each ws In ThisWorkbook.Worksheets
  ws.Activate

  Dim Ticker_Symbol As String

  Dim Total_Stock_Volume As Double
  Total_Stock_Volume = 0

  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
  
  Cells(1,9).Value = "Ticker"
  Cells(1,10).Value = "Total Stock Volume"
  Columns("I:J").EntireColumn.AutoFit

  lastrow = Cells(Rows.Count, 1).End(xlUp).Row
  
  For i = 2 To lastrow

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

      Ticker_Symbol = Cells(i, 1).Value

      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

      Range("I" & Summary_Table_Row).Value = Ticker_Symbol

      Range("J" & Summary_Table_Row).Value = Total_Stock_Volume
      
      Summary_Table_Row = Summary_Table_Row + 1

      Total_Stock_Volume = 0

    Else

      Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    End If

  Next i

ws.Cells(1, 1) = 1

Next

starting_ws.Activate 

End Sub
