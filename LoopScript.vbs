Attribute VB_Name = "Module1"
Sub ProcessStockData():
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentTicker As String
    Dim firstOpen As Double
    Dim lastClose As Double
    Dim totalVolume As Double
    Dim quarterlyChange As Double
    Dim percentChange As Double
    Dim i As Long
    Dim startRow As Long
    
    Set ws = ThisWorkbook.Sheets("Q1")
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    i = 2
    
    Do While i <= lastRow
        currentTicker = ws.Cells(i, 1).Value
        
        firstOpen = 0
        lastClose = 0
        totalVolume = 0
        startRow = i
        
        Do While we.Cells(i, 1).Value = currentTicker And i <= lastRow
            If firstOpen = 0 And ws.Cells(i, 3).Value > 0 Then
                firstOpen = ws.Cells(i, 3).Value
            End If
            
            lastClose = ws.Cells(i, 6).Value
            totalVolume = totalVolume + ws.Cells(i, 7).Value
            
            i = i + 1
       Loop
       
       If firstOpen > 0 Then
        quarterlyChange = lastClose - firstOpen
        percentChange = (quarterlyChange / firstOpen) * 100
    Else
        quarterlyChange = 0
        percentChange = 0
    End If
    
    ws.Cells(startRow, 8).Value = quarterlyChange
    ws.Cells(startRow, 9).Value = percentChange
    ws.Cells(startRow, 10).Value = totalVolume
    
    If quarterlyChnage > 0 Then
        ws.Cells(startRow, 8).Interior.Color = RGB(0, 255, 0) 'Green
    ElseIf quarterlyChange < 0 Then
        ws.Cells(startRow, 8).Interior.Color = RGB(255, 0, 0) 'Red
    Else
        ws.Cells(startRow, 8).Interior.ColorIndex = xlNone 'No color
    End If
    Loop
        
End Sub
