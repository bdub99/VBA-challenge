Attribute VB_Name = "Module1"
Sub multiyearstock()

' identify variable

For Each ws In Worksheets

Dim Ticker_Symbol As String
Dim Stock_Volume As Double
Dim Year_open, Year_close, Yearly_Change, Percent_Change As Currency
Dim Summary_Table_Row As LongLong


'declare

Ticker_Symbol = 0
Stock_Volume = 0
Yearly_Change = 0
Percent_Change = 0
Summary_Table_Row = 2


'loop for ticker symbol / find stock volume
    LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row
    For i = 2 To LastRow
        
        Ticker_Symbol = ws.Cells(i, 1).Value
        Stock_Volume = ws.Cells(i, 7).Value
        Year_open = ws.Cells(i, 3).Value
        Year_close = ws.Cells(i, 6).Value
        
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker_Symbol = ws.Cells(i, 1).Value
                
        Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
        
        ws.Range("I" & Summary_Table_Row).Value = Ticker_Symbol
        ws.Range("L" & Summary_Table_Row).Value = Stock_Volume
        ws.Range("J" & Summary_Table_Row).Value = Year_close - Year_open
        ws.Range("K" & Summary_Table_Row).Value = (Year_close - Year_open) / Year_open
        
        Summary_Table_Row = Summary_Table_Row + 1
        Ticker_Symbol = 0
        Stock_Volume = 0
        

        Else
            Stock_Volume = Stock_Volume + ws.Cells(i, 7).Value
            
     End If
        
    Next i
    For i = 2 To LastRow
        If Cells(i, 10).Value > 0 Then
                Cells(i, 10).Interior.ColorIndex = 4
            Else
                Cells(i, 10).Interior.ColorIndex = 3
        End If
       
    Next i
       
Next ws


End Sub

