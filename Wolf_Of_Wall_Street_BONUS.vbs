sub WolfOfWallStreetBONUS()

' Setting my variables
dim largestIncrease as Double
dim largestDecrease as Double
dim largestTV as Double
dim lastrow as Long
dim dataRange as Range
dim i as Integer
dim x as Integer
dim y as Integer
dim ws as Worksheet

For Each ws In Worksheets
' Finding the last row of data range
lastrow = ws.Cells(Rows.count, "I").End(xlUp).Row

' Setting range of primary data table
Set dataRange = ws.Range("I1:L" & lastrow)

' Calculating largest % increase
largestIncrease = Application.worksheetfunction.max(ws.Range("J:J"))/100
' Calculating largest % decrease
largestDecrease = Application.worksheetfunction.min(ws.Range("J:J"))/100
' Calculating the greatest total stock value
largestTV = Application.worksheetfunction.max(ws.Range("L:L"))

' Create headers of summary table
ws.Cells(1,16).value = "Ticker"
ws.Cells(1,17).value = "Value"
ws.Cells(2,15).value = "Greatest % Increase"
ws.Cells(3,15).value = "Greatest % Decrease"
ws.Cells(4,15).value = "Greatest Total Value"

ws.Cells(2,17).value = FormatPercent(largestIncrease,2)
for i = 2 to lastrow
    If (ws.Range("J" & i).value)/100 = largestIncrease Then
    ws.Cells(2,16).value = ws.Range("I" & i)
    End If
next i

ws.Cells(3,17).value = FormatPercent(largestDecrease,2)
for x = 2 to lastrow
    If (ws.Range("J" & x).value)/100 = largestDecrease Then
    ws.Cells(3,16).value = ws.Range("I" & x)
    End If
next x

ws.Cells(4,17).value = largestTV
for y = 2 to lastrow
    If ws.Range("L" & y).value = largestTV Then
    ws.Cells(4,16).value = ws.Range("L" & y).offset(0,-3)
    End If
next y

Next ws

End sub