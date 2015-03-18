Sub WorksheetLoop()

   Dim WS_Count As Integer
   Dim I As Integer
   Dim repName, repEmail As String

   ' Set WS_Count equal to the number of worksheets in the active
   ' workbook.
   WS_Count = ActiveWorkbook.Worksheets.Count

   ' Begin the loop.
   For I = 3 To WS_Count
        
      repName = Worksheets(2).Cells(I - 2, 1).Value
      repEmail = Worksheets(2).Cells(I - 2, 2).Value
      
      MsgBox ActiveWorkbook.Worksheets(repName).Name & " " & repEmail

      Worksheets(repName).Activate
      EmailReps repName, repEmail
      ' Insert your code here.
      ' The following line shows how to reference a sheet within
      ' the loop by displaying the worksheet name in a dialog box.
      

   Next I

End Sub
