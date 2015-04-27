Sub WorksheetLoop()

Dim WS_Count As Integer
Dim I, msgBoxResponse As Integer
Dim msgBoxPrompt, msgBoxTitle,repName, repEmail As String

' Set WS_Count equal to the number of worksheets in the active
' workbook.
WS_Count = ActiveWorkbook.Worksheets.Count

' Begin the loop.
For I = 3 To WS_Count

repName = Worksheets(2).Cells(I - 2, 1).Value
repEmail = Worksheets(2).Cells(I - 2, 2).Value

if msgBoxResponse <> vbCancel Then
  'the message box for verifying the email and worksheet'
  msgBoxPrompt = "Email:" & repEmail & vbNewLine & _
  "Sheet Name: " & ActiveWorkbook.Worksheets(I).Name & vbNewLine & _
  "Select Yes if the email and sheet are correct.  Select No if it is incorrect.  Select Cancel if you want to ignore this message for all future emails this run."
  msgBoxTitle = "Check if correct"
  msgBoxResponse = msgBox(msgBoxPrompt, vbYesNoCancel, msgBoxTitle)

End if
Worksheets(I).Activate
EmailReps repName, repEmail
' Insert your code here.
' The following line shows how to reference a sheet within
' the loop by displaying the worksheet name in a dialog box.
'MsgBox ActiveWorkbook.Worksheets(I).Name

Next I

End Sub
