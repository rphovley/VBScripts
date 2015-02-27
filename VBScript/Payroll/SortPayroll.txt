Sub sortPayroll()
Dim i, j As Integer
Dim sort As Worksheet
Dim sheetName, repName As String

sheetName = InputBox("What is the payroll sheet's name?", "Sheet Name")

Sheets.Add.Name = "New"
Set sort = Sheets("New")
i = 2
j = 2


With Sheets(sheetName)
Dim commisionSum, expenseSum, totalSum As Currency
repName = .Cells(2, 1).Value
Do Until IsEmpty(.Cells(i, 1))

If .Cells(i, 1).Value = repName Then

If .Cells(i, 5).Value = "Commission" Then
commisionSum = commisionSum + .Cells(i, 4).Value
End If

If .Cells(i, 5).Value = "Expense" Then
expenseSum = expenseSum + .Cells(i, 4).Value
End If

totalSum = totalSum + .Cells(i, 4).Value
Else

If totalSum > 0 Then
sort.Cells(j, 1).Font.Bold = True
sort.Cells(j, 1).Value = repName
sort.Cells(j, 2).Font.Bold = True
sort.Cells(j, 2).Value = totalSum

If commisionSum <> 0 Then
sort.Cells(j + 1, 1).Value = "Commision"
sort.Cells(j + 1, 2).Value = commisionSum
j = j + 1
End If

If expenseSum <> 0 Then
sort.Cells(j + 1, 1).Value = "Expense"
sort.Cells(j + 1, 2).Value = expenseSum
j = j + 1
End If


j = j + 1
End If
totalSum = 0
commisionSum = 0
expenseSum = 0
If .Cells(i, 5).Value = "Commission" Then
commisionSum = commisionSum + .Cells(i, 4).Value
End If

If .Cells(i, 5).Value = "Expense" Then
expenseSum = expenseSum + .Cells(i, 4).Value
End If

totalSum = totalSum + .Cells(i, 4).Value

End If
repName = .Cells(i, 1).Value
i = i + 1
Loop
End With


End Sub
