Attribute VB_Name = "Module1"

Sub Summary()

'Loop for all worksheets in the book
Dim ws As Worksheet
    
' Loop through each worksheet in the workbook
For Each ws In Worksheets

    'Create and format ticker summary chart header
    ws.Cells(1, 9).Value = "ticker"
    ws.Cells(1, 10).Value = "quarterly Change"
    ws.Cells(1, 11).Value = "percent Change"
    ws.Cells(1, 12).Value = "total Stock Volume"
        
    ws.Columns(10).NumberFormat = "0.00%" ' Percent format
        
    ws.Range("I:L").Columns.AutoFit ' AutoFit columns I to L

    'Identify last row in the actual sheet
    
    Dim LastTicker As Long
 
    LastTicker = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'Define needed variables for ticker summary chart

    Dim OpenQVal As Double
    Dim CloseQVal As Double
    Dim QChange As Double
    Dim PerChange As Double
    Dim TickerVol As Double
    Dim TickerRow As Integer
    
    OpenQVal = ws.Range("C2").Value
    CloseQVal = 0
    QChange = 0
    PerChange = 0
    TickerVol = 0
    TickerRow = 2
       
    'Summarize by ticker identifying when the ticker changes - Loop
    
    Dim i As Long
    
    For i = 2 To LastTicker

        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value) Then
        
        ws.Range("I" & TickerRow).Value = ws.Cells(i, 1).Value
        
        'Quarterly Change & Percentage Change calculation
        
        CloseQVal = ws.Cells(i, 6).Value
        QChange = CloseQVal - OpenQVal
        ws.Range("J" & TickerRow).Value = QChange
        
        PerChange = QChange / OpenQVal
        ws.Range("K" & TickerRow).Value = PerChange
        
        OpenQVal = ws.Cells(i + 1, 3).Value
        
        'Close ticker volume calculation
        
        TickerVol = TickerVol + ws.Cells(i, 7).Value
        ws.Range("L" & TickerRow).Value = TickerVol
        TickerRow = TickerRow + 1
        TickerVol = 0
            
        Else
        
        TickerVol = TickerVol + ws.Cells(i, 7).Value
            
        End If
    
    Next i
    
    'Color formatting for Quarterly Change - Loop
     
    Dim LQChange As Long
 
    LQChange = ws.Cells(Rows.Count, 10).End(xlUp).Row
    
    Dim QCi As Integer
    
    For QCi = 2 To LQChange
    
        If (ws.Range("J" & QCi).Value > 0) Then
    
        ws.Range("J" & QCi).Interior.ColorIndex = 4
    
        ElseIf (ws.Range("J" & QCi).Value < 0) Then
    
        ws.Range("J" & QCi).Interior.ColorIndex = 3
        
        End If
        
    Next QCi
    
    'Create and format ticker performance summary chart header and rows
    
    ws.Range("O1").Value = "ticker"
    ws.Range("P1").Value = "value"
    ws.Range("N2").Value = "Greatest % Increase"
    ws.Range("N3").Value = "Greatest % Decrease"
    ws.Range("N4").Value = "Greatest Total Volume"
    
    ws.Range("P2:P3").NumberFormat = "0.00%"
    
    'Define needed variables for performance summary chart
    
    Dim GreatestIncTicker As String
    Dim GreatestInc As Double
    
    Dim GreatestDecTicker As String
    Dim GreatestDec As Double
    
    Dim GreatestVolTicker As String
    Dim GreatestVol As Double
    
    'Find Greatest % Increase and Greatest % Decrease - Loop
    
    Dim PCi As Integer
        
    GreatestInc = ws.Range("K2").Value
    GreatestDec = ws.Range("K2").Value
    GreatestVol = 0
    
    For PCi = 3 To LQChange
    
        If (ws.Range("K" & PCi).Value > GreatestInc) Then
        
        GreatestInc = ws.Range("K" & PCi).Value
        GreatestIncTicker = ws.Range("I" & PCi).Value
        
        ElseIf (ws.Range("K" & PCi).Value < GreatestDec) Then
        
        GreatestDec = ws.Range("K" & PCi).Value
        GreatestDecTicker = ws.Range("I" & PCi).Value
        
        End If
        
    Next PCi
    
    ws.Range("O2").Value = GreatestIncTicker
    ws.Range("P2").Value = GreatestInc
    
    ws.Range("O3").Value = GreatestDecTicker
    ws.Range("P3").Value = GreatestDec
    
    'Find Greatest Total Volume - Loop
    
    Dim TVi As Integer
        
    GreatestVol = ws.Range("L2").Value
    
    For TVi = 3 To LQChange
    
        If (ws.Range("L" & TVi).Value > GreatestVol) Then
        
        GreatestVol = ws.Range("L" & TVi).Value
        GreatestVolTicker = ws.Range("I" & TVi).Value
        
        End If
        
    Next TVi
    
    ws.Range("O4").Value = GreatestVolTicker
    ws.Range("P4").Value = GreatestVol
    
    'Autofit the ticker performance summary chart after adding the data
    
    Dim ColumnsTickerPerformance As Range
    
    Set ColumnsTickerPerformance = ws.Range("N:P")
    
    ColumnsTickerPerformance.Columns.AutoFit
    
Next ws

End Sub
