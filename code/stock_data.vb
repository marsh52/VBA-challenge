Sub stock_data()
    ' count number of sheets in the book
    Dim sheets As Integer
    sheets = ThisWorkbook.Worksheets.Count

    Dim currentRow As Long, rows As Long, i As Long, quarterOpenPrice As Double, quarterClosePrice As Double, quarterChange As Double, totalStockVolume As LongLong, currentTicker As String, outputTableRow As Long, percentChange As Double
    Dim highPercent As Double, lowPercent As Double, highVolume As LongLong
    Dim highPercSymbol As String, lowPercSymbol As String, highVolSymbol As String
    Dim ws As Worksheet
    
    ' iterate though sheets
    For i = 1 To sheets
    
        currentRow = 2
        rows = 0
        totalStockVolume = 0
        outputTableRow = 2
        highPercent = 0
        lowPercent = 0
        highVolume = 0
        Set ws = ThisWorkbook.sheets(i)
        
        ' function that populates output table headers
        Call headers(ws, 1)
        
        ' this while loop iterates through all cells in the column containing data using the isempty function
        While (Not IsEmpty(ws.Cells(currentRow, 1)))
            ' first row conditional
            If (currentRow = 2) Then
                currentTicker = ws.Cells(currentRow, 1).Value
                quarterOpenPrice = ws.Cells(currentRow, 3).Value
                totalStockVolume = ws.Cells(currentRow, 7).Value
                
            ' change in ticker
            ElseIf (ws.Cells(currentRow, 1).Value <> currentTicker) Then
                ' update different variables
                ws.Cells(outputTableRow, 9).Value = currentTicker
                currentTicker = ws.Cells(currentRow, 1).Value
                quarterClosePrice = ws.Cells(currentRow - 1, 6).Value
                percentChange = (quarterClosePrice / quarterOpenPrice) - 1
                quarterChange = quarterClosePrice - quarterOpenPrice
                ' update output table
                ws.Cells(outputTableRow, 10).Value = quarterChange
                ws.Cells(outputTableRow, 11).Value = percentChange
                ws.Cells(outputTableRow, 11).NumberFormat = "0.00%"
                ws.Cells(outputTableRow, 12).Value = totalStockVolume
                ' change color of cell
                Call colorChange(ws, percentChange, outputTableRow)
                
                ' check if percent change is high/low
                If (percentChange < lowPercent) Then
                    lowPercent = percentChange
                    lowPercSymbol = ws.Cells(outputTableRow, 9).Value
                ElseIf (percentChange > highPercent) Then
                    highPercent = percentChange
                    highPercSymbol = ws.Cells(outputTableRow, 9).Value
                End If
                
                ' check if stock volume is high
                If (totalStockVolume > highVolume) Then
                    highVolume = totalStockVolume
                    highVolSymbol = ws.Cells(outputTableRow, 9).Value
                End If
                
                ' update values for next ticker
                totalStockVolume = ws.Cells(currentRow, 7).Value
                quarterOpenPrice = ws.Cells(currentRow, 3).Value
                ' increase row on output table
                outputTableRow = outputTableRow + 1
                quarterClosePrice = 0
                
            ' last row
            ElseIf (IsEmpty(ws.Cells(currentRow + 1, 1).Value)) Then
                ' this is all very similar to the change in ticker conditional
                ws.Cells(outputTableRow, 9).Value = currentTicker
                quarterClosePrice = ws.Cells(currentRow, 6)
                percentChange = (quarterClosePrice / quarterOpenPrice) - 1
                quarterChange = quarterClosePrice - quarterOpenPrice
                totalStockVolume = totalStockVolume + ws.Cells(currentRow, 7).Value
                ws.Cells(outputTableRow, 10).Value = quarterChange
                ws.Cells(outputTableRow, 11).Value = percentChange
                ws.Cells(outputTableRow, 11).NumberFormat = "0.00%"
                ws.Cells(outputTableRow, 12).Value = totalStockVolume
                
                Call colorChange(ws, percentChange, outputTableRow)
                
                If (percentChange < lowPercent) Then
                    lowPercent = percentChange
                    lowPercSymbol = ws.Cells(outputTableRow, 9).Value
                ElseIf (percentChange > highPercent) Then
                    highPercent = percentChange
                    highPercSymbol = ws.Cells(outputTableRow, 9).Value
                End If
                ' check if stock volume is high
                If (totalStockVolume > highVolume) Then
                    highVolume = totalStockVolume
                    highVolSymbol = ws.Cells(outputTableRow, 9).Value
                End If
                
                Call summaryTable(ws, lowPercent, highPercent, highVolume, lowPercSymbol, highPercSymbol, highVolSymbol)
            ' sum totalStockVolume for rows without conditions
            ElseIf (ws.Cells(currentRow, 1).Value = currentTicker) Then
                totalStockVolume = totalStockVolume + ws.Cells(currentRow, 7).Value
            End If
            rows = rows + 1
            currentRow = currentRow + 1
        Wend
    Next i
End Sub

Sub headers(ws As Worksheet, row As Integer)
    Dim headerList(3) As String
    Dim header As Variant
    Dim column As Integer
    headerList(0) = "Ticker"
    headerList(1) = "Quarterly Change"
    headerList(2) = "Percent Change"
    headerList(3) = "Total Stock Volume"
    column = 9
    ' i need to put these ^ in the table
    For Each header In headerList
        ws.Cells(row, column).Value = header
        column = column + 1
    Next header
    
End Sub

Sub summaryTable(ws As Worksheet, lowPercent As Double, highPercent As Double, highVolume As LongLong, lowPercSymbol As String, highPercSymbol As String, highVolSymbol As String)
    Dim headerList(1) As String
    Dim rowNames(2) As String
    Dim newTableRow As Integer
    Dim column As Integer
    newTableRow = 1
    column = 15
    
    headerList(0) = "Ticker"
    headerList(1) = "Value"
    rowNames(0) = "Greatest % Increase"
    rowNames(1) = "Greatest % Decrease"
    rowNames(2) = "Greatest Total Volume"
    
    For Each header In headerList
        ws.Cells(newTableRow, column + 1).Value = header
        column = column + 1
    Next header
    
    For Each Name In rowNames
        ws.Cells(newTableRow + 1, 15).Value = Name
        newTableRow = newTableRow + 1
    Next Name
    
    For i = 2 To 4
        If (i = 2) Then
            ws.Cells(i, 16).Value = highPercSymbol
            ws.Cells(i, 17).Value = highPercent
            ws.Cells(i, 17).NumberFormat = "0.00%"
        ElseIf (i = 3) Then
            ws.Cells(i, 16).Value = lowPercSymbol
            ws.Cells(i, 17).Value = lowPercent
            ws.Cells(i, 17).NumberFormat = "0.00%"
        ElseIf (i = 4) Then
            ws.Cells(i, 16).Value = highVolSymbol
            ws.Cells(i, 17).Value = highVolume
        End If
    Next i
    
End Sub

Sub colorChange(ws As Worksheet, percentChange As Double, outputTableRow As Long)
    If (percentChange > 0) Then
        ws.Cells(outputTableRow, 11).Interior.ColorIndex = 4
    ElseIf (percentChange < 0) Then
        ws.Cells(outputTableRow, 11).Interior.ColorIndex = 3
    End If
End Sub
