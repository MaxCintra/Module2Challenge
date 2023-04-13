Attribute VB_Name = "Module1"
Sub stock_data()

For Each ws In Sheets

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

Dim Total As LongLong
Total = 0

Dim year_open As Double
Dim year_close As Double
Dim year_change As Double
Dim percent_change

ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

ws.Cells(2, 15).Value = "Greatest % Increase"
ws.Cells(3, 15).Value = "Greatest % Decrease"
ws.Cells(4, 15).Value = "Greatest Total Volume"
ws.Cells(1, 16).Value = "Ticker"
ws.Cells(1, 17).Value = "Value"

Dim Summary_Table_Row As Integer
Summary_Table_Row = 1

For i = 2 To lastrow
'''''''''''''''''''''''''''''''''''''''''
    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        
        ticker = ws.Cells(i, 1)
        Total = Total + ws.Cells(i, 7)
        Summary_Table_Row = Summary_Table_Row + 1
        ws.Range("I" & Summary_Table_Row).Value = ticker
        ws.Range("L" & Summary_Table_Row).Value = Total
    
        Total = 0
        
    Else
    
        Total = Total + ws.Cells(i, 7).Value
        
    End If
'''''''''''''''''''''''''''''''''''
    If Right(ws.Cells(i, 2), 4) = "0102" Then
        year_open = ws.Cells(i, 3).Value
    ElseIf Right(ws.Cells(i, 2), 4) = "1231" Then
        year_close = ws.Cells(i, 6).Value
        year_change = year_close - year_open
        percent_change = year_change / year_open
        ws.Range("J" & Summary_Table_Row).Value = year_change
        ws.Range("J" & Summary_Table_Row).NumberFormat = "0.00"
        ws.Range("K" & Summary_Table_Row).Value = percent_change
        ws.Range("K" & Summary_Table_Row).NumberFormat = "0.00%"
        
        
    End If
''''''''''''''''''''''''''''''''''''''
    If ws.Cells(Summary_Table_Row, "J").Value > 0 Then
        ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 4
    ElseIf ws.Cells(Summary_Table_Row, "J").Value < 0 Then
        ws.Cells(Summary_Table_Row, "J").Interior.ColorIndex = 3
    
    End If
''''''''''''''''''''''''''''''''''''''

Next i
''''''''''''''''''''''''''''''''''''''''
Dim percrng As Range
Dim stockrng As Range
Dim percmax As Double
Dim percmin As Double
Dim stockmax As Double
Dim ticker1 As String
Dim ticker2 As String
Dim ticker3 As String
Set percrng = ws.Range("K:K")
Set stockrng = ws.Range("L:L")
percmax = Application.WorksheetFunction.Max(percrng)
percmin = Application.WorksheetFunction.Min(percrng)
stockmax = Application.WorksheetFunction.Max(stockrng)
ticker1 = Application.Match(percmax, percrng, 0)
ticker2 = Application.Match(percmin, percrng, 0)
ticker3 = Application.Match(stockmax, stockrng, 0)
ws.Cells(2, 17).Value = percmax
ws.Cells(3, 17).Value = percmin
ws.Cells(4, 17).Value = stockmax
ws.Cells(2, 16).Value = ws.Cells(ticker1, "I").Value
ws.Cells(3, 16).Value = ws.Cells(ticker2, "I").Value
ws.Cells(4, 16).Value = ws.Cells(ticker3, "I").Value
ws.Range("Q2", "Q3").NumberFormat = "0.00%"

'''''''''''''''''''''''''''''''''''''''''


ws.Columns("A:Q").AutoFit


Next ws


End Sub
