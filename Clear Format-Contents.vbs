
Sub Clear()
Dim ws As Worksheet

For Each ws In Worksheets
    ' Code to clear contents of summary data table. This will make it faster to check to see if my code is working.
    
    ws.Range("I:L").ClearContents
    ws.Range("I:L").ClearFormats
    ws.Range("O:Q").ClearContents
    ws.Range("O:Q").ClearFormats

Next ws

End Sub


