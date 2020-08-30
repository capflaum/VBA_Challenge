# VBA_Challenge
# VBA Homework #2

Sub VBA_challenge()
    Dim TickerName As String
    TickerName = 0
    Dim Row As Integer
    Row = 2
    
    Dim OpenValue As Double
    Dim CloseValue As Double
    Dim PercentageChange As Double
    OpenValue = Cells(2, 3).Value
    
    Dim StartStockRow As Long
    StartStockRow = 2
    
#Loop with Row counter
RowCount = Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To RowCount

#Ticker List
    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        TickerName = Cells(i, 1).Value
        Range("I" & Row).Value = TickerName
        
#Total Stock Volume
        Range("L" & Row).Value = Application.Sum(Range(Cells(StartStockRow, 7), Cells(i, 7)))
#Yearly Change (Open Value before loop to get first value)
        CloseValue = Cells(i, 6).Value
        
        Range("J" & Row).Value = (CloseValue - OpenValue)
   
#Percentage Change with conditional formatting
        PercentageChange = (CloseValue - OpenValue) / OpenValue
        Range("K" & Row).Value = FormatPercent(PercentageChange)
        
        If PercentageChange < 0 Then
            Cells(Row, 11).Interior.ColorIndex = 3
        Else
            Cells(Row, 11).Interior.ColorIndex = 4
        End If
#Set up next Loop
        Row = Row + 1
        StartStockRow = i + 1
        OpenValue = Cells(i + 1, 3).Value
        
    
    End If
    
Next i

End Sub

  
