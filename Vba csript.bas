Attribute VB_Name = "Module1"
Sub Script()

    Dim i As Long
    Dim lastrow As Long
    Dim ws As Worksheet
    
    Dim openingPrice As Double
    Dim closingPrice As Double
    Dim quarterlyChange As Double
    Dim currentTicker As String
    Dim nextTicker As String
    Dim PercentChange As Double
    Dim totalVolume As Double
    Dim outputRow As Long

' Variables to store greatest changes and total volume
    Dim greatestIncrease As Double
    Dim greatestDecrease As Double
    Dim greatestVolume As Double
    Dim tickerIncrease As String
    Dim tickerDecrease As String
    Dim tickerVolume As String

    ' Initialize variables
    greatestIncrease = 0
    greatestDecrease = 0
    greatestVolume = 0



    ' Loop through all sheets
    For Each ws In Worksheets
    
        ' Determine the Last Row
        lastrow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
        ' Add a Column for the Ticker
        ws.Range("I1").EntireColumn.Insert
        
        ' Add column headers
        ws.Cells(1, 9).Value = "Ticker"
        ws.Cells(1, 10).Value = "Quarterly Change"
        ws.Cells(1, 11).Value = "Percent Change"
        ws.Cells(1, 12).Value = "Total Stock Volume"
        
        outputRow = 2  ' Start from the second row for output data
        totalVolume = 0  ' Initialize total volume

        ' Loop through all the stock prices
        For i = 2 To lastrow
            
            currentTicker = ws.Cells(i, 1).Value
            nextTicker = ws.Cells(i + 1, 1).Value  ' Get the ticker for the next row
            
            ' Add the current volume to total
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            ' Check if we are at the start of a new ticker
            If ws.Cells(i - 1, 1).Value <> currentTicker Then
                ' Take the opening price of the new ticker
                openingPrice = ws.Cells(i, 3).Value
            End If
            
            
            ' Check if we are at the end of the current ticker
            If currentTicker <> nextTicker Or i = lastrow Then
                ' Take the closing price of the current ticker
                closingPrice = ws.Cells(i, 6).Value
                
                ' Add the ticker to column I (output row)
                ws.Cells(outputRow, 9).Value = currentTicker
                
                ' Output the total volume to column L
                ws.Cells(outputRow, 12).Value = totalVolume
                
                ' Reset the total volume for the next ticker
                totalVolume = 0
                
                ' Calculate the quarterly change
                quarterlyChange = closingPrice - openingPrice
                
                ' Output the quarterly change to column J
                ws.Cells(outputRow, 10).Value = quarterlyChange
                
                
                
                ' Check if the quarterly change is positive or negative and color accordingly
            If quarterlyChange > 0 Then
                ws.Cells(outputRow, 10).Interior.ColorIndex = 4   ' Green
                
           ElseIf quarterlyChange < 0 Then
                ws.Cells(outputRow, 10).Interior.ColorIndex = 3   ' Red
                
           Else
                ws.Cells(outputRow, 10).Interior.ColorIndex = 0   ' No color for zero change
                
           End If
                
                
                
                ' Calculate the percentage change
                If openingPrice <> 0 Then
                    PercentChange = ((closingPrice - openingPrice) / openingPrice)
                Else
                    PercentChange = 0
                End If
                
                ' Output the percent change to column K
                ws.Cells(outputRow, 11).Value = PercentChange
                
                
                ' Compare the percentage changes and total volume between current and previous tickers
                If PercentChange > greatestIncrease Then
                    greatestIncrease = PercentChange
                    tickerIncrease = currentTicker
                End If
                
                If PercentChange < greatestDecrease Then
                    greatestDecrease = PercentChange
                    tickerDecrease = currentTicker
                End If
                
                If ws.Cells(outputRow, 12).Value > greatestVolume Then
                    greatestVolume = ws.Cells(outputRow, 12).Value
                    tickerVolume = currentTicker
                End If
                
                ' Move to the next output row
                outputRow = outputRow + 1
            End If
        Next i
        
  
    
    
    ' Display the results in a new set of columns
    Dim resultRow As Long
    resultRow = 2 ' Start writing results in row 2

    ' Add headers for results
    
    ws.Cells(1, 16).Value = "Ticker"
    ws.Cells(1, 17).Value = "Value"
    
    ' Output greatest increase
    ws.Cells(resultRow, 15).Value = "Greatest % Increase"
    ws.Cells(resultRow, 16).Value = tickerIncrease
    ws.Cells(resultRow, 17).Value = Format(greatestIncrease, "0.00") & "%"

    ' Output greatest decrease
    resultRow = resultRow + 1
    ws.Cells(resultRow, 15).Value = "Greatest % Decrease"
    ws.Cells(resultRow, 16).Value = tickerDecrease
    ws.Cells(resultRow, 17).Value = Format(greatestDecrease, "0.00") & "%"

    ' Output greatest total volume
    resultRow = resultRow + 1
    ws.Cells(resultRow, 15).Value = "Greatest Total Volume"
    ws.Cells(resultRow, 16).Value = tickerVolume
    ws.Cells(resultRow, 17).Value = greatestVolume

Next ws




End Sub

