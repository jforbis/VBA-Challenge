Sub ticker()

' Declaring variables
Dim i As Long
Dim ticker As String
Dim yearlyChange As Double
Dim volume As Long
Dim lastrow As Long
Dim dataTableRow As Long

dataTableRow = 2

' Creating headers for new data table
Cells(1, 9).Value = "Ticker"
Cells(1, 10).Value = "Yearly Change"
Cells(1, 11).Value = "Percent Change"
Cells(1, 12).Value = "Total Stock Volume"

' Create table listing Tickers and corresponding data
 
 'Loop through all stock data on ONE sheet
 For i = 2 To 60000

    If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        ticker = Cells(i, 1).Value
        Range("I" & dataTableRow).Value = ticker
        dataTableRow = dataTableRow + 1
    End If

Next i

End Sub