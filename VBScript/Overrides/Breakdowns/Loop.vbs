Sub WorksheetLoop()

   Dim WS_Count As Integer
   Dim I As Integer
   Dim repName, repEmail As String

   ' Set WS_Count equal to the number of worksheets in the active
   ' workbook.
   WS_Count = Worksheets("Reps").Cells(2, 1).End(xlDown).Row

   ' Begin the loop.
   For I = 2 To WS_Count + 1
        
      repName = Worksheets("Reps").Cells(I, 1).Value
      repEmail = Worksheets("Reps").Cells(I, 2).Value
      
      MsgBox ActiveWorkbook.Worksheets(repName).Name & " " & repEmail

      Worksheets(repName).Activate
      EmailReps repName, repEmail
      ' Insert your code here.
      ' The following line shows how to reference a sheet within
      ' the loop by displaying the worksheet name in a dialog box.
      

   Next I

End Sub

