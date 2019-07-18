Sub accumStockVolume():

'
'
'  author - Emily Mo
'  date - Feb 2 2019
'  boot camp U Miami = Data Analytics
'

Dim ws As Worksheet
Dim openPrice As Double
Dim yrEndClosePrice As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim totalVolume As Double
Dim i As Long
Dim lastRow As Long

Dim biggestInc As Double
Dim biggestDec As Double
Dim biggestVol As Double
Dim biggestIncT As String
Dim biggestDecT As String
Dim biggestVolT As String
Dim x As Long
Dim notFound As Boolean
Dim currentTicker As String



For Each ws In Worksheets



'
'
' Greatest % increase - biggestInc and biggestIncT
' Greatest % Decrease - biggestDec and biggestDecT
' Greatest total volume - biggestVol and biggestVolT
'
'
'
    ws.Range("I1").Value = "Ticker"
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").Value = "Percent Change"
    ws.Range("L1").Value = "Total Stock Volume"
    
    biggestInc = 0
    biggestDec = 0
    biggestVol = 0

    
'
'   x is the index of the output ticker row
'   initialize totalVolume and currentTicker - info for the currently processed ticker
'
    
    x = 2
    totalVolume = 0
    currentTicker = " "

    
    lastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row

    For i = 2 To lastRow

'
'   in case stock not opened on first day but in the middle of the year,
'   need to loop until the first day of business (open / close prices not = 0)
'
'
        If (ws.Cells(i, 3) = 0 And ws.Cells(i, 6) = 0) Then
            notFound = True
            Do While notFound
                i = i + 1
                If (ws.Cells(i, 3) <> 0 Or ws.Cells(i, 6) <> 0) Then
                    notFound = False
                End If
            Loop
        End If
    
        If ws.Cells(i, 1) <> currentTicker Then
        
            openPrice = ws.Cells(i, 3).Value
            currentTicker = ws.Cells(i, 1)
            
        End If
        
        
        totalVolume = totalVolume + ws.Cells(i, 7).Value
    
'
'   if there is a change of ticker or processing lastRow, then output the accumulation for the current ticker
'
    
        If (ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1)) Or (i = lastRow) Then
            yrEndClosePrice = ws.Cells(i, 6).Value
        
            yearlyChange = yrEndClosePrice - openPrice
'
'   if openPrice is 0, division will fall over
'
            If openPrice = 0 Then
                percentChange = 0
            Else
                percentChange = yearlyChange / openPrice
            End If
'
'   positive change - green or else red
'
    
            If yearlyChange >= 0 Then
                ws.Cells(x, 10).Interior.Color = vbGreen
            Else
                ws.Cells(x, 10).Interior.Color = vbRed
            End If
        
'
'   output the ticker being processed but before a change of ticker
'
        
        
            ws.Cells(x, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(x, 10).Value = yearlyChange
            ws.Cells(x, 11).NumberFormat = "0.00%"
            ws.Cells(x, 11).Value = percentChange
            ws.Cells(x, 12).Value = totalVolume
            x = x + 1
'
'   find the greatest inc/dec/total vol
'
        
        
            If totalVolume > biggestVol Then
                biggestVol = totalVolume
                biggestVolT = ws.Cells(i, 1).Value
            End If
        
            If percentChange > biggestInc Then
                biggestInc = percentChange
                biggestIncT = ws.Cells(i, 1).Value
            End If
            If percentChange < biggestDec Then
                biggestDec = percentChange
                biggestDecT = ws.Cells(i, 1).Value
            End If
        
'
'   reset total volume for the next ticker
'
        
            totalVolume = 0
        
        End If
        
    Next i
    


'
'   at end of worksheet, output greatest inc/dec/total vol
'

    ws.Range("P1").Value = "Ticker"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Increase"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("O4").Value = "Greatest Total Volume"

    ws.Range("P2").Value = biggestIncT
    ws.Range("P3").Value = biggestDecT
    ws.Range("P4").Value = biggestVolT
    ws.Range("Q2").NumberFormat = "0.00%"
    ws.Range("Q3").NumberFormat = "0.00%"
    ws.Range("Q2").Value = biggestInc
    ws.Range("Q3").Value = biggestDec
    ws.Range("Q4").Value = biggestVol

    ws.Columns("A:Q").AutoFit

Next

End Sub