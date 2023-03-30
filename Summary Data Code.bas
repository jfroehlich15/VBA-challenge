Attribute VB_Name = "Module1"
Sub SummaryData()

'create loop for every worksheet in the workbook

Dim ws As Worksheet

For Each ws In Worksheets

'create column headers
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"


'create i variable as Long
Dim i As Long


'create ticker variable as string
Dim ticker As String

'create open variable as double and set value
Dim start As Double
start = 0

'create close variable as double and set value
Dim finish As Double
finish = 0

'create Yearly change variable as double and set value
Dim year_change As Double
year_change = 0

'create percent change variable and set value
Dim percent_change As Double
percent_change = 0


'create volume variable as long and set value
Dim volume As LongLong
volume = 0

'create integer for summary table and set value
Dim Summ_Table As Integer
Summ_Table = 2

'code to find last row in each worksheet
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row + 1

'set open starting total
start = ws.Range("C2").Value

'create for loop
For i = 2 To LastRow

'check if ticker is the same and if not then
If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

'set ticker name
ticker = ws.Cells(i, 1).Value

'set close total
finish = ws.Cells(i, 6).Value

'set yearly change
year_change = finish - start

'set percent change
percent_change = year_change / start

'set total stock volume
volume = volume + ws.Cells(i, 7).Value

'Print ticker name
ws.Range("I" & Summ_Table).Value = ticker

'print yearly change and format
ws.Range("J" & Summ_Table).Value = year_change
ws.Range("J" & Summ_Table).NumberFormat = "0.00"

'print percent change and format
ws.Range("K" & Summ_Table).Value = percent_change
ws.Range("K" & Summ_Table).NumberFormat = "0.00%"

'print total volume and format
ws.Range("L" & Summ_Table).Value = volume
ws.Columns("L").NumberFormat = "0"

'add 1 to summ table to move to next row
Summ_Table = Summ_Table + 1

'reset variables to 0 and start to the open price of the next ticker
start = ws.Cells(i + 1, 3).Value
finish = 0
year_change = 0
volume = 0


'if they are the same ticker
Else

'add to volume variable and continue
volume = volume + ws.Cells(i, 7).Value

End If

Next i

'-------------------------------
'new loop to format cells with color once all rows have been created
For i = 2 To LastRow

'if greater than 0 format green
If ws.Cells(i, 10) > 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 4

'if less than 0 format red
ElseIf ws.Cells(i, 10) < 0 Then
ws.Cells(i, 10).Interior.ColorIndex = 3

'if equal to 0 format as normal cell
Else
ws.Cells(i, 10).Interior.ColorIndex = 0

End If

Next i


'format width of cells at the very end
ws.Columns("I:L").AutoFit



Next


End Sub



