Sub WolfOfWallStreetSequel()

' Declaring variables
Dim i As Long
Dim a as Long
Dim ticker As String
Dim last As Double
Dim yearlyChange As Double
Dim percentChange As Double
Dim volume As Double
Dim rangeCount As Long
Dim lastrow As Long
Dim dataTable As Long
Dim lastDataTableRow As Long
Dim ws As Worksheet

' Loop everything over all worksheets inside the workbook
For Each ws In Worksheets

'Finding last row for main table and 2nd data table that is created.
lastrow = ws.Cells(Rows.count, "A").End(xlUp).Row

' Creating headers for new data table
dataTable = 2
ws.Cells(1, 9).Value = "Ticker"
ws.Cells(1, 10).Value = "Yearly Change"
ws.Cells(1, 11).Value = "Percent Change"
ws.Cells(1, 12).Value = "Total Stock Volume"

 'Loop through all stock data on ONE sheet
 For i = 2 To lastrow
    ' Setting up a Counter to be used to count length of ticker range
    If ws.Cells(i, 1).Value = ws.Cells(i + 1, 1) Then
        rangeCount = rangeCount + 1

        'Stock Volume total per ticker range
        volume = volume + ws.Cells(i, 7).Value
        
    ElseIf ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        'List out each Ticker type
        ticker = ws.Cells(i, 1).Value
        ws.Range("I" & dataTable).Value = ticker

        'Variables for start and end of ticker range
        start = ws.Cells(i - rangeCount, 3).Value
        last = ws.Cells(i, 6).Value

        If start <> 0 Then
        'List out yearly change from open to close
        yearlyChange = last - start
        'List out percent change from open to close
        percentChange = yearlyChange / last
        Else
        percentChange = 0
        yearlyChange = 0
        End If
        
        ws.Range("J" & dataTable).Value = yearlyChange


        ws.Range("K" & dataTable).Value = FormatPercent(percentChange, 2)
    
        'Sum total stock volume per ticker
        volume = volume + ws.Cells(i, 7).Value
        ws.Range("L" & dataTable).Value = volume

        'Reset values
        rangeCount = 0
        volume = 0

        dataTable = dataTable + 1

    End If

Next i

lastDataTableRow = ws.Cells(Rows.count, "J").End(xlUp).Row

for a = 2 to lastDataTableRow
    ' Conditional to change cell color for positive or negative change
    If ws.Cells(a,10).value > 0 Then
        ' Color Code
        ws.Cells(a,10).Interior.ColorIndex = 4

    ElseIf ws.Cells(a,10).value < 0 Then
        ' Color Code
        ws.Cells(a,10).Interior.ColorIndex = 3

    End If

Next a

' Go to next worksheet
Next ws

End Sub

