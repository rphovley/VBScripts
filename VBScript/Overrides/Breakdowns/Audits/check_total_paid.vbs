Sub checkTotalPaid()
Dim rep_count, bottomPrintRow, topPrintRow As Integer
   Dim I As Integer
   Dim totalDiff As Currency

   Dim repCol, kWCol, rateCol, totalCol, valueCol, diffCol, checkCol As Integer
       repCol   = 2
       kwCol    = 4
       rateCol  = 5
       totalCol = 6
       valueCol = 12
       diffCol  = 13
       checkCol = 14
   Dim repName, repEmail As String

   ' Set rep_count equal to the number of worksheets in the active
   ' workbook.
   rep_count = ThisWorkbook.Worksheets("Reps").Cells(2, 1).End(xlDown).Row

   ' Begin the loop.
   For I = 2 To rep_count
      totalDiff = 0
      repName = ThisWorkbook.Worksheets("Reps").Cells(I, 1).Value
      repEmail = ThisWorkbook.Worksheets("Reps").Cells(I, 2).Value

      
      ' Insert your code here.
      ' The following line shows how to reference a sheet within
      ' the loop by displaying the worksheet name in a dialog box.
      With ThisWorkbook.Worksheets(repName)
      	.Activate
      	topPrintRow    = .Cells(3, 2).End(xlDown).End(xlDown).Row + 2
      	bottomPrintRow = .Cells(3, 2).End(xlDown).End(xlDown).End(xlDown).Row
      	
      	.Cells(topPrintRow - 1, valueCol) = "Job Value"
      	.Cells(topPrintRow- 1, diffCol)  = "Difference"
      	For row = topPrintRow To bottomPrintRow

      		.Cells(row, valueCol).Value = .Cells(row, kWCol).Value * .Cells(row, rateCol).Value
      		If .Cells(row, valueCol) < .Cells(row, totalCol) Then
      			.Cells(row, checkCol).Value = "CHECK"
      			.Cells(row, diffCol).Value = .Cells(row, totalCol) -.Cells(row, valueCol)
      			totalDiff = totalDiff + .Cells(row, diffCol).Value
      		Else
      			.Cells(row, checkCol).Value = "GOOD"
      		End If
      	Next row
      	'sort data by check col'
      	.Range("B" & topPrintRow & ":N" & bottomPrintRow).Sort key1:=Range("N" & topPrintRow & ":N" & bottomPrintRow), _
        order1:=xlAscending, Header:=xlNo
      	'Adjusts column width of the new tab
	    .Range("C1:N1").EntireColumn.AutoFit
      End With

      ThisWorkbook.Worksheets("Reps").Cells(I, 3).Value = totalDiff
   Next I
End Sub