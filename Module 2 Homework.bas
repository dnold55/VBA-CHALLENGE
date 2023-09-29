Attribute VB_Name = "Module1"
Sub RunAnalysisOnAllSheets()

    Dim sheet As Worksheet
    
    For Each sheet In ThisWorkbook.Sheets
        Call StockAnalysis(sheet)
    Next sheet

End Sub

Sub StockAnalysis(ByVal ws As Worksheet)

    ws.Cells(1, 9).Value = "Ticker"
    ws.Cells(1, 10).Value = "Yearly Change"
    ws.Cells(1, 11).Value = "Percent Change"
    ws.Cells(1, 12).Value = "Total Volume"
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    ws.Cells(2, 15).Value = "Greatest % Increase"
    ws.Cells(3, 15).Value = "Greatest % Decrease"
    ws.Cells(4, 15).Value = "Greatest Total Volume"

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Dim ticker As String
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim totalVolume As Double
    
    Dim outputRow As Long
    outputRow = 2
    
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String
    
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    
    Dim i As Long
    For i = 2 To lastRow
        
        ticker = ws.Cells(i, 1).Value
        
        If ws.Cells(i - 1, 1).Value <> ticker Then
           
            openingPrice = ws.Cells(i, 3).Value
            totalVolume = 0
        End If
        
        totalVolume = totalVolume + ws.Cells(i, 7).Value
        
        If ws.Cells(i + 1, 1).Value <> ticker Then
            
            closingPrice = ws.Cells(i, 6).Value
            
            
            ws.Cells(outputRow, 9).Value = ticker
            ws.Cells(outputRow, 10).Value = closingPrice - openingPrice
            ws.Cells(outputRow, 11).Value = (closingPrice - openingPrice) / openingPrice
            ws.Cells(outputRow, 12).Value = totalVolume
            
            Dim percentChange As Double
            percentChange = (closingPrice - openingPrice) / openingPrice
            If percentChange > greatestIncrease Then
                greatestIncrease = percentChange
                tickerIncrease = ticker
            ElseIf percentChange < greatestDecrease Then
                greatestDecrease = percentChange
                tickerDecrease = ticker
            End If
            If totalVolume > greatestVolume Then
                greatestVolume = totalVolume
                tickerVolume = ticker
            End If
            
            
            
            If ws.Cells(outputRow, 10).Value > 0 Then
                ws.Cells(outputRow, 10).Interior.Color = RGB(0, 255, 0)
            Else
                ws.Cells(outputRow, 10).Interior.Color = RGB(255, 0, 0)
            End If
            ws.Cells(outputRow, 11).NumberFormat = "0.00%"
            outputRow = outputRow + 1
        End If
        
    Next i
    
     outputRow = outputRow + 1
    ws.Cells(outputRow, 10).Value = "Greatest % Increase"
    ws.Cells(outputRow, 11).Value = tickerIncrease
    ws.Cells(outputRow, 12).Value = greatestIncrease
    ws.Cells(outputRow, 12).NumberFormat = "0.00%"
    
    outputRow = outputRow + 1
    ws.Cells(outputRow, 10).Value = "Greatest % Decrease"
    ws.Cells(outputRow, 11).Value = tickerDecrease
    ws.Cells(outputRow, 12).Value = greatestDecrease
    ws.Cells(outputRow, 12).NumberFormat = "0.00%"
    
    outputRow = outputRow + 1
    ws.Cells(outputRow, 10).Value = "Greatest Total Volume"
    ws.Cells(outputRow, 11).Value = tickerVolume
    ws.Cells(outputRow, 12).Value = greatestVolume

End Sub
