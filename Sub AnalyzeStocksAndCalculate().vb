Sub AnalyzeStocksAndCalculate()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim summaryRow As Long
    Dim i As Long
    Dim ticker As String
    Dim totalVolume As Double
    Dim openPrice As Double
    Dim closePrice As Double
    Dim change As Double
    Dim percentChange As Double
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim greatestIncreaseTicker As String
    Dim greatestDecreaseTicker As String
    Dim greatestVolumeTicker As String
    
    'Initialize variables for greatest calculations across all sheets
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0
    
    ' Loop through all worksheets in the workbook
    For Each ws In ThisWorkbook.Worksheets
        
        ' Find the last row of data based on the first column
        lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        ' Create headers for new columns in summary table
        ws.Cells(1, 8).Value = "Ticker"
        ws.Cells(1, 9).Value = "Total Stock Volume"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        
        ' Initialize summary row
        summaryRow = 2

        ' Initialize variables for the first ticker
        ticker = ws.Cells(2, 1).Value
        openPrice = ws.Cells(2, 3).Value
        totalVolume = 0

        ' Loop through rows to calculate changes and total volume
        For i = 2 To lastRow
        
            ' Accumulate volume
            totalVolume = totalVolume + ws.Cells(i, 7).Value

            ' Check if the ticker symbol changes or if it's the last row
            If ws.Cells(i + 1, 1).Value <> ticker Or i = lastRow Then
                ' Close price is from the current row
                closePrice = ws.Cells(i, 6).Value
                
                ' Calculate quarterly change and percent change
                change = closePrice - openPrice
                If openPrice <> 0 Then
                    percentChange = (change / openPrice)
                Else
                    percentChange = 0
                End If

                ' Write calculations to summary table
                ws.Cells(summaryRow, 8).Value = ticker
                ws.Cells(summaryRow, 9).Value = totalVolume
                ws.Cells(summaryRow, 10).Value = change
                ws.Cells(summaryRow, 11).Value = percentChange
                
                ' Apply conditional formatting to the Quarterly Change column
                If change > 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = vbGreen
                ElseIf change < 0 Then
                    ws.Cells(summaryRow, 10).Interior.Color = vbRed
                End If
                
                ' Check for greatest increase, decrease, and total volume
                If percentChange > greatestIncrease Then
                    greatestIncrease = percentChange
                    greatestIncreaseTicker = ticker
                End If
                If percentChange < greatestDecrease Then
                    greatestDecrease = percentChange
                    greatestDecreaseTicker = ticker
                End If
                If totalVolume > greatestVolume Then
                    greatestVolume = totalVolume
                    greatestVolumeTicker = ticker
                End If

                ' Move to the next summary row
                summaryRow = summaryRow + 1

                ' Reset variables for the new ticker
                ticker = ws.Cells(i + 1, 1).Value
                openPrice = ws.Cells(i + 1, 3).Value
                totalVolume = 0
            End If
        Next i
        
        ' Format Percent Change column as percentage
        ws.Range(ws.Cells(2, 11), ws.Cells(summaryRow - 1, 11)).NumberFormat = "0.00%"

        ' Format Quarterly Change column
        ws.Range(ws.Cells(2, 10), ws.Cells(summaryRow - 1, 10)).NumberFormat = "0.00"

        ' Format Total Stock Volume with comma separators
        ws.Range(ws.Cells(2, 9), ws.Cells(summaryRow - 1, 9)).NumberFormat = "#,##0"

    Next ws

    ' Output the greatest values
    MsgBox "Greatest % Increase: " & greatestIncreaseTicker & " - " & Format(greatestIncrease, "0.00%") & vbCrLf & _
           "Greatest % Decrease: " & greatestDecreaseTicker & " - " & Format(greatestDecrease, "0.00%") & vbCrLf & _
           "Greatest Total Volume: " & greatestVolumeTicker & " - " & Format(greatestVolume, "#,##0")
           
End Sub

