Sub calculate()
Dim Worksheet As Worksheet
'Loop through all Sheets
For s = 1 To ThisWorkbook.Sheets.Count

Set Worksheet = ThisWorkbook.Sheets(s)

With Worksheet
    printrow = 2
    currentSymbol = .Cells(2, 1).Value
    openingPrice = .Cells(2, 3).Value
    closingPrice = .Cells(2, 6).Value
    totalVol = .Cells(2, 7).Value
    lastRow = .Cells(1, 1).End(xlDown).Row
    'Set up Headers
    .Cells(1, 9).Value = "Ticker"
    .Cells(1, 10).Value = "Yearly Change"
    .Cells(1, 11).Value = "Percent Change"
    .Cells(1, 12).Value = "Total Volume"
    'Challenge 1
    .Cells(1, 15).Value = "Ticker"
    .Cells(1, 16).Value = "Value"
    .Cells(2, 14).Value = "Greatest % increase"
    .Cells(3, 14).Value = "Greatest % decrease"
    .Cells(4, 14).Value = "Greatest total volume"
     
     increaseSymbol = currentSymbol
     increaseValue = 0
     decreaseSymbol = currentSymbol
     decreaseValue = 0
     volumeSymbol = currentSymbol
     volumeValue = 0
     
     
    
        For currentrow = 3 To lastRow
            newSym = .Cells(currentrow, 1).Value
            If StrComp(currentSymbol, newSym, vbTextCompare) = 0 Then
                closingPrice = .Cells(currentrow, 6).Value
                totalVol = .Cells(currentrow, 7).Value + totalVol
            Else
                'Calculate change
                yearChange = closingPrice - openingPrice
                If openingPrice > 0 Then
                    perChange = yearChange / openingPrice
                Else
                    perChange = 0
                End If
                
                'Find the Greatest Percent
                If perChange > increaseValue Then
                    increaseValue = perChange
                    increaseSymbol = currentSymbol
                    
                End If
                If perChange < decreaseValue Then
                     decreaseValue = perChange
                     decreaseSymbol = currentSymbol
                    
                End If
                
                    If totalVol > volumeValue Then
                    volumeValue = totalVol
                    volumeSymbol = currentSymbol
                    End If
                             
                
                'Print Summary
                .Cells(printrow, 9).Value = currentSymbol
                .Cells(printrow, 10).Value = yearChange
                .Cells(printrow, 11).Value = perChange
                .Cells(printrow, 12).Value = totalVol
                'Advance printrow to the next line
                printrow = printrow + 1
                'Reset for new Symbol
                currentSymbol = .Cells(currentrow, 1).Value
                openingPrice = .Cells(currentrow, 3).Value
                closingPrice = .Cells(currentrow, 6).Value
                totalVol = .Cells(currentrow, 7).Value
    
            End If
    
    Next currentrow
    
    'Report Greatest Values
    .Cells(2, 15).Value = increaseSymbol
    .Cells(3, 15).Value = decreaseSymbol
    .Cells(4, 15).Value = volumeSymbol
    .Cells(2, 16).Value = increaseValue
    .Cells(3, 16).Value = decreaseValue
    .Cells(4, 16).Value = volumeValue
        
    'Conditional Formatting concatenation
    With .Range("j2:j" & printrow).FormatConditions _
        .Add(xlCellValue, xlGreater, 0)
        .Interior.ColorIndex = 4
    
    End With
    With .Range("j2:j" & printrow).FormatConditions _
        .Add(xlCellValue, xlLess, 0)
        .Interior.ColorIndex = 3
    
    End With
    'NumberFormat
    .Range("k2:K" & printrow).NumberFormat = "0.00%"
    .Range("p2:p3").NumberFormat = "0.00%"
    'NumberFormat0.00
    .Range("j2:j" & printrow).NumberFormat = "0.00"
End With
Next s
End Sub


