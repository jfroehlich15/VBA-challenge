Attribute VB_Name = "Module2"
Sub SummaryTable()

'create for loop for each worksheet

Dim ws As Worksheet

For Each ws In Worksheets

'code to find last row in each worksheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1

'-----------------------------
'create variables for greatest increase and define greatest increase as 0
Dim i As Long
Dim Greatest_Increase As Double
Dim Greatest_Increase2 As Double
Dim tikcer As String

Greatest_Increase = 0


'loop through each row in column K searching for the max percentage
For i = 2 To LastRow
Greatest_Increase2 = WorksheetFunction.Max(ws.Range("K" & i))
If Greatest_Increase2 > Greatest_Increase Then
Greatest_Increase = Greatest_Increase2
ticker = ws.Cells(i, 9).Value


End If
Next i

'add values to summary table location
ws.Range("P2").Value = ticker
ws.Range("Q2").Value = Greatest_Increase


'------------------------
'create variables for greatest decrease and define greatest decrease as 0

Dim Greatest_Decrease As Double
Dim Greatest_Decrease2 As Double
Dim decrease_ticker As String

Greatest_Decrease = 0


'loop through each row in column K searching for the min percentage
For i = 2 To LastRow
Greatest_Decrease2 = WorksheetFunction.Min(ws.Range("K" & i))
If Greatest_Decrease2 < Greatest_Decrease Then
Greatest_Decrease = Greatest_Decrease2
decrease_ticker = ws.Cells(i, 9).Value

End If
Next i

'add values to summary table location
ws.Range("P3").Value = decrease_ticker
ws.Range("Q3").Value = Greatest_Decrease


'------------------------
'create variables for greatest volume and define greatest volume as 0

Dim Greatest_Volume As LongLong
Dim Greatest_Volume2 As LongLong
Dim volume_ticker As String

Greatest_Volume = 0


'loop through each row in column L searching for the max volume
For i = 2 To LastRow
Greatest_Volume2 = WorksheetFunction.Max(ws.Range("L" & i))
If Greatest_Volume2 > Greatest_Volume Then
Greatest_Volume = Greatest_Volume2
volume_ticker = ws.Cells(i, 9).Value

End If
Next i


'add values to summary table location
ws.Range("P4").Value = volume_ticker
ws.Range("Q4").Value = Greatest_Volume

'-------------------------------------------

'format data type in summary table
ws.Range("Q2:Q3").NumberFormat = "0.00%"
ws.Range("Q4").NumberFormat = "0"


'create headers for summary table
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


'format width of summary table columns
ws.Columns("O").AutoFit
ws.Columns("P").ColumnWidth = 10
ws.Columns("Q").AutoFit

Next


End Sub


