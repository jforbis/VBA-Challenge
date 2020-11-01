sub Ticker()

' Creating headers for new data table
Cells(1,9).value = "Ticker"
Cells(1,10).value = "Yearly Change"
Cells(1,11).value = "Percent Change"
Cells(1,12).value = "Total Stock Volume"

' Create column listing Tickers
 Dim i as Integer
 Dim ticker as String
 Dim lastrow as Integer

lastrow = Cells(Rows.Count, 1). End(xlUp).Rows

 For i = 2 to lastrow

    If Cells(i+1,1).value<>cells(i1).value then
    msgbox(cells(i,1).value)

    End If

Next i

End sub