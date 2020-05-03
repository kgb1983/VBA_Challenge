Sub VBAHomework()

' Set variables

Dim Ticker As String
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim TotalVolume As Long

' Find lastrow

lastrow = Cells(Rows.Count, 1).End(xlUp).Row

' Create Summary Table

Dim Summary_Table_Row As Integer
Summary_Table_Row = 2

' Loop through each worksheet

Dim ws As Worksheet

For Each ws In Worksheets

' Loop through all tickers

For i = 2 To lastrow

If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then

' Set the Ticker Name
Ticker = Cells(i, 1).Value

' Add to the Volume Total

VolumeTotal = VolumeTotal + Cells(i, 7).Value

' Print the Ticker Name in the Summary Table
Range("H" & Summary_Table_Row).Value = Ticker

' Print the VolumeTotal in the Summary Table
Range("K" & Summary_Table_Row).Value = VolumeTotal

' Add one to the summary table row
Summary_Table_Row = Summary_Table_Row + 1

Else

' Add to the Volume Total

VolumeTotal = VolumeTotal + Cells(i, 7).Value

End If

Next i

Next ws

End Sub