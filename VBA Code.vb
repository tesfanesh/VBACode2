Sub testing_1()


'loop through all sheets
For Each ws In Worksheets
 
 ' Created a Variable to Hold no of rows ,ticker ,total,open value,close value , yearly change and percent cahnge
 Dim no_of_rows As Double, ticker As String, total As Double, open_value As Double, close_value As Double, yearly_change As Double
 Dim percent_change As Double

 'calculated the no of rows
 no_of_rows = ws.Cells(Rows.Count, 1).End(xlUp).Row

 'assign the headers to the cells
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Volume"
 'set the total to 0
 total = 0

 'assign the value to row from where we want to print our data
 Row = 2

 'set the open_value to the first value in open column
 open_value = ws.Cells(2, 3)

 'loop through the all rows
 For i = 2 To no_of_rows

 ' when the value of the next cell is different than that of the current cell
   If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then

      'set the value into Ticker column
       ws.Cells(Row, 9).Value = ws.Cells(i, 1).Value

       'Add to the Ticker Total
       total = ws.Cells(i, 7).Value + total
       ws.Cells(Row, 12).Value = total

       'set the last row of particular ticker into close value
       close_value = ws.Cells(i, 6).Value

       'the difference between close value and open value is our yearly change
       yearly_change = (close_value - open_value)

       'set yearly_change value to yearly change column
       ws.Cells(Row, 10).Value = yearly_change

       'check the condiotion on open value and close value
       If (open_value = 0 And close_value = 0) Then
          percent_change = 0
       ElseIf (open_value = 0 And close_value <> 0) Then
          percent_change = 1
       Else
          percent_change = yearly_change / open_value
          ws.Cells(Row, 11).Value = percent_change
          ws.Cells(Row, 11).NumberFormat = "0.00%"
       End If
       
       'set the next row value of open column as open_value
       open_value = ws.Cells(i + 1, 3).Value

       'increase the row to store the value in next row
       Row = Row + 1

       'set total = 0 for next ticker
       total = 0

   'if both rows are same just add the total to total value
   Else
      total = total + ws.Cells(i, 7).Value
   End If
Next i


'calculate the last row for formatting the cells based on their values
yearly_LastRow = ws.Cells(Rows.Count, 9).End(xlUp).Row

'loop through all rows
For j = 2 To yearly_LastRow

   'if value greater than or equal to 0 set the color of the cell green
   If (ws.Cells(j, 10).Value > 0 Or ws.Cells(j, 10).Value = 0) Then
                ws.Cells(j, 10).Interior.ColorIndex = 10

   'f not set the color of the cells red
   ElseIf ws.Cells(j, 10).Value < 0 Then
                ws.Cells(j, 10).Interior.ColorIndex = 3
   End If
Next j


'assign the headers to the cells
ws.Range("o2").Value = "Greatest % Increase"
ws.Range("o3").Value = "Greatest % Decrease"
ws.Range("o4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"


'find  range for finding maximum and minimum of yearly change and maximum of total volume
Set myRange = ws.Range("K2:K" & yearly_LastRow)
Set myRange1 = ws.Range("L2:L" & yearly_LastRow)


'loop throgh all rows
For Z = 2 To yearly_LastRow

  'check for maximum yearly change
  If ws.Cells(Z, 11).Value = ws.Application.WorksheetFunction.Max(myRange) Then

   'assign maximum value to the assigned cell
   ws.Range("P2").Value = ws.Cells(Z, 9).Value
   ws.Range("q2").Value = ws.Cells(Z, 11).Value
   
   'change the format to percentage
   ws.Range("q2").NumberFormat = "0.00%"
   
   'check for minimum yearly change
  ElseIf (ws.Cells(Z, 11).Value = ws.Application.WorksheetFunction.Min(myRange)) Then

   'assign maximum value to the assigned cell
   ws.Range("P3").Value = ws.Cells(Z, 9).Value
   ws.Range("q3").Value = ws.Cells(Z, 11).Value

   'change the format to percentage
   ws.Range("q3").NumberFormat = "0.00%"

  'check for maximum total volume
  ElseIf (ws.Cells(Z, 12).Value = ws.Application.WorksheetFunction.Max(myRange1)) Then

   'assign maximum value to the assigned cell
   ws.Range("P4").Value = ws.Cells(Z, 9).Value
   ws.Range("q4").Value = ws.Cells(Z, 12).Value
  End If

'Next row
Next Z
'next sheet
Next ws
End Sub
