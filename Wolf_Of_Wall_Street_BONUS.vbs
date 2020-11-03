sub WolfOfWallStreetBONUS()

' Setting my variables
dim largestIncrease as Double
dim largestDecrease as Double
dim largestTV as Double
dim lastrow as Long
dim dataRange as Range
dim i as Integer

' Finding the last row of data range
lastrow = Cells(Rows.count, "I").End(xlUp).Row

' Setting range of primary data table
Set dataRange = Range("I1:L" & lastrow)

' Calculating largest % increase
largestIncrease = Application.worksheetfunction.max(Range("J:J"))
' Calculating largest % decrease
largestDecrease = Application.worksheetfunction.min(Range("J:J"))
' Calculating the greatest total stock value
largestTV = Application.worksheetfunction.max(Range("L:L"))

' Create headers of summary table
Cells(1,16).value = "Ticker"
Cells(1,17).value = "Value"
Cells(2,15).value = "Greatest % Increase"
Cells(3,15).value = "Greatest % Decrease"
Cells(4,15).value = "Greatest Total Value"

Cells(2,17).value = largestIncrease
for i = 2 to lastrow
    If Range("J" & i).value = largestIncrease Then
    Cells(2,16).value = Range("J" & i).offset(0,-1)
    End If
next i

Cells(3,17).value = largestDecrease
for i = 2 to lastrow
    If Range("J" & i).value = largestDecrease Then
    Cells(3,16).value = Range("J" & i).offset(0,-1)
    End If
next i

Cells(4,17).value = largestTV
for i = 2 to lastrow
    If Range("L" & i).value = largestTV Then
    Cells(4,16).value = Range("L" & i).offset(0,-3)
    End If
next i

End sub