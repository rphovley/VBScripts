Sub createNatesEvo(ByVal workBookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
	Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
	    createdDateCol, repEmailCol, isDocSignedCol, isFinalContractCol As Integer

	    customerCol        = 1   
	    jobCol             = 2
	    kWCol              = 3
	    statusCol          = 4
	    subStatusCol       = 5
	    createdDateCol     = 6
	    isDocSignedCol     = 7
	    isFinalContractCol = 8
	    repEmailCol        = 9
	    
	Dim masCustomerCol, masJobCol, masKWCol, masStatusCol, masSubStatusCol, _
		masDateCol, masRepEmailCol As Integer

		masCustomerCol        = 1   
	    masJobCol             = 2
	    masKWCol              = 3
	    masStatusCol          = 4
	    masSubStatusCol       = 5
	    masCreatedDateCol     = 7
	    masRepEmailCol        = 17

	Dim nateIsDocSignedCol, nateIsFinalContractCol As Integer

		nateIsDocSignedCol     = 2   
	    nateIsFinalContractCol = 3

	''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsm"       
        Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

	''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
		Dim inputDataSheet, printSheet, masterInput, natesSheet As Worksheet
		Set inputDataSheet = NatesEvolution.Worksheets("Master Input")
		Set printSheet     = NatesEvolution.Worksheets("Nate's Evolution")
		Set masterInput    = NatesEvolution.Worksheets("Master Input")
		Set natesSheet     = NatesEvolution.Worksheets("Docs Signed Input")

	''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
		Dim inputRow, printRow, natesJobRow As Integer
			printRow = 2

	'''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
		Dim inputData() As cJobData
		Dim currentRep  As cJobData
		Dim printRep    As cJobData

	'''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
		Dim inputDataSize As Long
			inputDataSize = inputDataSheet.UsedRange.Rows.Count - 1
			ReDim inputData(inputDataSize)

	'''''''''Found Job in Nate's Input''''''''''''''''''''''''''''''
		Dim isJobFound As Boolean
	'''''''''''Date constant to consider for new jobs''''''''''''''
   	Const NEWJOBDATE = #11/30/2014#

	''''''''''''''''''TURN OFF SCREEN UPDATING''''''''''''''''''''''
	application.screenupdating = False
	'''''''''''''''''''''''''Print out'''''''''''''''''''''''''''''
		For inputRow = 2 To inputDataSize + 1
			isJobFound = True

			''''''Determine if it is part of new pay structure'''''''

			If masterInput.Cells(inputRow, masCreatedDateCol).Value >= NEWJOBDATE Then
				'''''''''''''''''''Find job in Nate's Sheet'''''''''''''''
				On Error GoTo jobIdNotFound:
				natesJobRow = Application.WorksheetFunction.Match(masterInput.Cells(inputRow, masJobCol).Value, natesSheet.Range("E:E"), 0)
				'natesJobrow = "=INDEX($B:$B, MATCH(" + Col_Letter(masJobCol) + CStr(inputRow) + ",$E:$E,0))"



				'''''''''''''''''''Print out the Output'''''''''''''''''''
				With printSheet
					''''''From Master Report'''''
					.Cells(inputRow, customerCol).Value     = masterInput.Cells(inputRow, masCustomerCol).Value
					.Cells(inputRow, jobCol).Value          = masterInput.Cells(inputRow, masJobCol).Value
					.Cells(inputRow, kWCol).Value           = masterInput.Cells(inputRow, masKWCol).Value
					.Cells(inputRow, statusCol).Value       = masterInput.Cells(inputRow, masStatusCol).Value
					.Cells(inputRow, subStatusCol).Value    = masterInput.Cells(inputRow, masSubStatusCol).Value
					.Cells(inputRow, createdDateCol).Value  = masterInput.Cells(inputRow, masCreatedDateCol).Value				
					.Cells(inputRow, repEmailCol).Value     = masterInput.Cells(inputRow, masRepEmailCol).Value

					''From Nate's Input'''
					if isJobFound Then
						.Cells(inputRow, isDocSignedCol).Value     = natesSheet.Cells(natesJobRow, nateIsDocSignedCol).Value
						'.Cells(inputRow, isDocSignedCol).Value     = "=INDEX($B:$B, MATCH(" + Col_Letter(CStr(masJobCol)) + CStr(inputRow) + ",$E:$E,0))"
						.Cells(inputRow, isFinalContractCol).Value = natesSheet.Cells(natesJobrow, nateIsFinalContractCol).Value
						'.Cells(inputRow, isFinalContractCol).Value = "=INDEX($C:$C, MATCH(" + Col_Letter(CStr(masJobCol)) + CStr(inputRow) + ",$E:$E,0))"
					Else
						.Cells(inputRow, isDocSignedCol).Value     = "N"
						.Cells(inputRow, isFinalContractCol).Value = "N"
					End If

				End With
			End If

		Next inputRow
''''''''''''''''''TURN ON SCREEN UPDATING''''''''''''''''''''''
		application.screenupdating = False
		jobIdNotFound:
			isJobFound = False
			Resume Next
End Sub
