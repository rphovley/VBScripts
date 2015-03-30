Private Sub FirstName()
	'David McRitchie 2000-03-23 programming
	'http://www.geocities.com/davemcritchie/excel/join.htm#firstname
	'Put cells in range from "LastName, FirstName" to "FirstName LastName"
	Application.Calculation = xlManual
	Dim cell As Range
	Dim cPos As Long
	For Each cell In Selection.SpecialCells(xlConstants, xlTextValues)
	cPos = InStr(1, cell, ",")
	If cPos > 1 Then
	origcell = cell.Value
	cell.Value = Trim(Mid(cell, cPos + 1)) & " " _
	& Trim(Left(cell, cPos - 1))
	End If

	Next cell
	Application.Calculation = xlAutomatic 'xlCalculationAutomatic
End Sub