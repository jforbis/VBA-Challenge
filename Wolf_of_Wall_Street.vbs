Sub stocks()

' Declaring variables
Dim i As Long
Dim ticker As String
Dim yearlyChange As Double
Dim percentChange as Double
Dim volume As Long
Dim rangeCount as Long
Dim lastrow As Long
Dim dataTable As Long

lastrow = Cells(Rows.Count, "A").End(xlUp).Row

' Creating headers for new data table
dataTable = 2
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Create table listing Tickers and corresponding data
 
 'Loop through all stock data on ONE sheet
 For i = 2 To lastrow
    ' Setting up a Counter to be used to count length of ticker range
    If Cells(i,1).value = Cells(i+1,1) Then
        rangeCount = rangeCount + 1
        
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'List out each Ticker type
        ticker = Cells(i, 1).Value
        Range("I" & dataTable).Value = ticker

        'Variables for start and end of ticker range
        start = cells(i-rangeCount,3).value
        last = cells(i,6).value       
        
        'List out yearly change from open to close 
        yearlyChange = last-start
        Range("J" & dataTable).Value = yearlyChange 
      
        'List out percent change from open to close
        percentChange = yearlyChange/last
        Range("K" & dataTable).value = FormatPercent(percentChange)
    
        'Sum total stock volume per ticker
        volume = volume + Cells(i,7).value
        Range("L" & dataTable).value = volume

        rangeCount = 0

        dataTable = dataTable + 1

    End If

Next i

End Sub