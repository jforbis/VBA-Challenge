Sub WolfOfWallStreet()

' Declaring variables
Dim i As Long
Dim ticker As String
Dim yearlyChange As Double
Dim percentChange As Double
Dim volume As Double
Dim rangeCount As Long
Dim lastrow As Long
Dim dataTable As Long
Dim lastDataTableRow As Long

lastrow = Cells(Rows.count, "A").End(xlUp).Row
lastDataTableRow = Cells(Rows.count, "J").End(xlUp).Row

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
    If Cells(i, 1).Value = Cells(i + 1, 1) Then
        rangeCount = rangeCount + 1

        'Stock Volume total per ticker range
        volume = volume + Cells(i, 7).Value
        
    ElseIf Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        'List out each Ticker type
        ticker = Cells(i, 1).Value
        Range("I" & dataTable).Value = ticker

        'Variables for start and end of ticker range
        Start = Cells(i - rangeCount, 3).Value
        last = Cells(i, 6).Value
        
        'List out yearly change from open to close
        yearlyChange = last - Start

        Range("J" & dataTable).Value = yearlyChange

        'List out percent change from open to close
        percentChange = yearlyChange / last
        Range("K" & dataTable).Value = FormatPercent(percentChange)
    
        'Sum total stock volume per ticker
        volume = volume + Cells(i, 7).Value
        Range("L" & dataTable).Value = volume

        'Reset values
        rangeCount = 0
        volume = 0

        dataTable = dataTable + 1

    End if

    ' Conditional to change cell color for positive or negative change
    IF yearlyChange > 0 Then   
        ' Color Code
        Cells(10, i).Interior.ColorIndex = 4

    Else 
        ' Color Code
        Cells(10, i).Interior.ColorIndex = 3

    End If  
Next i

End Sub
