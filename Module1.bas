Attribute VB_Name = "Module1"

Sub AnalyzeStockData()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim yearlyChange As Double
    Dim percentChange As Double
    Dim totalVolume As Double
    Dim maxIncrease As Double
    Dim maxDecrease As Double
    Dim maxVolume As Double
    Dim maxIncreaseTicker As String
    Dim maxDecreaseTicker As String
    Dim maxVolumeTicker As String
    Dim outputRow As Long
    
    ' Initialize variables for tracking output row
    outputRow = 2
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        ' Find the last row with data in column A
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        ' Initialize variables for each worksheet
        openingPrice = ws.Cells(2, 3).Value
        closingPrice = ws.Cells(lastRow, 6).Value
        totalVolume = Application.WorksheetFunction.Sum(ws.Range("G2:G" & lastRow))
        
        ' Calculate yearly change and percent change
        yearlyChange = closingPrice - openingPrice
        If openingPrice <> 0 Then
            percentChange = (yearlyChange / openingPrice) * 100
        Else
            percentChange = 0
        End If
        
        ' Output results
        ws.Cells(outputRow, 14).Value = "Ticker"
        ws.Cells(outputRow, 15).Value = "Yearly Change"
        ws.Cells(outputRow, 16).Value = "Percent Change"
        ws.Cells(outputRow, 17).Value = "Total Stock Volume"
        ws.Cells(outputRow + 1, 14).Value = ws.Cells(2, 1).Value
        ws.Cells(outputRow + 1, 15).Value = yearlyChange
        ws.Cells(outputRow + 1, 16).Value = percentChange
        ws.Cells(outputRow + 1, 17).Value = totalVolume
        
        
        ' Find and apply conditional formatting for yearly change
        If yearlyChange > 0 Then
            ws.Cells(outputRow & 1 + 10).Interior.ColorIndex = 4
        Else
            ws.Cells(outputRow & 1 + 10).Interior.ColorIndex = 3
        End If
        
        
        
        ' Find the greatest % increase, % decrease, and total volume
        If percentChange > maxIncrease Then
            maxIncrease = percentChange
            maxIncreaseTicker = ws.Cells(2, 1).Value
        End If
        If percentChange < maxDecrease Then
            maxDecrease = percentChange
            maxDecreaseTicker = ws.Cells(2, 1).Value
        End If
        If totalVolume > maxVolume Then
            maxVolume = totalVolume
            maxVolumeTicker = ws.Cells(2, 1).Value
            End If
        
        ' Move to the next output row
        outputRow = outputRow + 2
    Next ws
    
    ' Output the greatest % increase, % decrease, and total volume in the last worksheet
    Dim summaryWs As Worksheet
    Set summaryWs = ThisWorkbook.Worksheets.Add
    summaryWs.Name = "Summary"
    summaryWs.Cells(1, 1).Value = "Greatest % Increase"
    summaryWs.Cells(2, 1).Value = "Greatest % Decrease"
    summaryWs.Cells(3, 1).Value = "Greatest Total Volume"
    summaryWs.Cells(1, 2).Value = maxIncreaseTicker
    summaryWs.Cells(2, 2).Value = maxDecreaseTicker
    summaryWs.Cells(3, 2).Value = maxVolumeTicker
    summaryWs.Cells(1, 3).Value = maxIncrease
    summaryWs.Cells(2, 3).Value = maxDecrease
    summaryWs.Cells(3, 3).Value = maxVolume
End Sub


