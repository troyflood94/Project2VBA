Attribute VB_Name = "Module1"
Sub summarizePage()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim ticker As Variant
    Dim tickerRow As Long
    Dim openPrice As Double
    Dim closePrice As Variant
    Dim openCloseDiff As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim i As Long
    Dim firstOpenPrice As Double
    Dim lastClosePrice As Double
    Dim isFirstTicker As Boolean ' Flag to track the first ticker
    Dim greatestIncreaseTicker As Variant
    Dim greatestDecreaseTicker As Variant
    Dim highestVolumeTicker As Variant
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim highestVolume As Double

    ' Loop through each worksheet
    For Each ws In ThisWorkbook.Worksheets
        ' Reset tickerRow counter
        tickerRow = 2
        
        ' Find the last used row
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Set first open price
        firstOpenPrice = 0
        
        ' Set flag for the first ticker
        isFirstTicker = True
        
        ' Loop through each row in column A
        For i = 2 To lastRow
            ' Check if the current ticker .Value is different from the previous row
            If ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
                ' If not the first ticker, assign the last close price
                If Not isFirstTicker Then
                    closePrice = ws.Cells(i - 1, 6).Value
                    lastClosePrice = IIf(IsNumeric(closePrice), CDbl(closePrice), 0)
                    
                    ' Calculate open-close difference
                    openCloseDiff = lastClosePrice - firstOpenPrice
                    
                    ' Calculate percent change
                    If firstOpenPrice <> 0 Then
                        percentChange = (openCloseDiff / firstOpenPrice) * 100
                    Else
                        percentChange = 0
                    End If
                    
                    ' Calculate total volume
                    totalVolume = Application.WorksheetFunction.SumIf(ws.Range("A:A"), ticker, ws.Range("G:G"))
                    
                    ' Print ticker and summary in column J on respective page
                    ws.Cells(tickerRow, "J").Value = ticker
                    ws.Cells(tickerRow, "K").Value = openCloseDiff
                    ws.Cells(tickerRow, "L").Value = percentChange
                    ws.Cells(tickerRow, "M").Value = totalVolume
                    
                    ' Conditional formatting
                    If openCloseDiff >= 0 Then
                        ws.Cells(tickerRow, "K").Interior.Color = RGB(0, 255, 0) ' Green
                    Else
                        ws.Cells(tickerRow, "K").Interior.Color = RGB(255, 0, 0) ' Red
                    End If
                    
                    ' Update the values for greatest increase, greatest decrease, and highest volume
                    If percentChange > greatestIncrease Then
                        greatestIncrease = percentChange
                        greatestIncreaseTicker = ticker
                    End If
                    
                    If percentChange < greatestDecrease Then
                        greatestDecrease = percentChange
                        greatestDecreaseTicker = ticker
                    End If
                    
                    If totalVolume > highestVolume Then
                        highestVolume = totalVolume
                        highestVolumeTicker = ticker
                    End If
                    
                    ' Move to the next row for the next ticker
                    tickerRow = tickerRow + 1
                End If
                
                ' Assign the ticker value
                ticker = ws.Cells(i, 1).Value
                
                ' Get the first open price for the new ticker
                firstOpenPrice = ws.Cells(i, 3).Value
                
                ' Update the flag to indicate that it's not the first ticker anymore
                isFirstTicker = False
            End If
        Next i
        
        ' Label and create values for the greatest percent increase, greatest percent decrease, and highest total volume
        
        ws.Range("P1").Value = "Greatest % Increase"
        ws.Range("Q1").Value = greatestIncreaseTicker
        ws.Range("R1").Value = greatestIncrease
        
        ws.Range("P2").Value = "Greatest % Decrease"
        ws.Range("Q2").Value = greatestDecreaseTicker
        ws.Range("R2").Value = greatestDecrease
        
        ws.Range("P3").Value = "Highest Total Volume"
        ws.Range("Q3").Value = highestVolumeTicker
        ws.Range("R3").Value = highestVolume
    Next ws
End Sub






