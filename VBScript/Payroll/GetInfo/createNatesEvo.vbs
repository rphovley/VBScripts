Sub createNatesEvo(ByVal workBookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
	Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
	    createdDateCol, repEmailCol, finalContratCol, stateCol As Integer

	    customerCol        = 1   
	    jobCol             = 2
	    kWCol              = 3
	    statusCol          = 4
	    subStatusCol       = 5
	    createdDateCol     = 6
	    repEmailCol        = 7
	    finalContratCol    = 8
	    stateCol           = 9
	    
	Dim masCustomerCol, masJobCol, masKWCol, masStatusCol, masSubStatusCol, _
		masDateCol, masFinalCol, masRepEmailCol, masStateCol As Integer

		masCustomerCol        = 1   
	    masJobCol             = 2
	    masKWCol              = 3
	    masStatusCol          = 4
	    masSubStatusCol       = 5
	    masCreatedDateCol     = 7
	    masFinalCol           = 8
	    masRepEmailCol        = 17
	    masStateCol           = 18

	Dim nateIsDocSignedCol, nateIsFinalContractCol As Integer

		nateIsDocSignedCol     = 2   
	    nateIsFinalContractCol = 3

	''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsm"       
        Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

	''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
		Dim jobDataSheet, printSheet, masterInput, natesSheet As Worksheet
		Set jobDataSheet = NatesEvolution.Worksheets("Master Input")
		Set printSheet     = NatesEvolution.Worksheets("Nate's Evolution")
		Set masterInput    = NatesEvolution.Worksheets("Master Input")

	''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
		Dim inputRow, printRow, natesJobRow As Integer
			printRow = 2

	'''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
		Dim jobDataSize As Long
			jobDataSize = jobDataSheet.UsedRange.Rows.Count - 1
			ReDim jobData(jobDataSize)
	'''''''''''Date constant to consider for new jobs''''''''''''''
   	Const NEWJOBDATE = #11/30/2014#

	
	'''''''''''''''''''''''''Print out'''''''''''''''''''''''''''''
		For inputRow = 2 To jobDataSize + 1
			isJobFound = True

			''''''Determine if it is part of new pay structure'''''''

			If masterInput.Cells(inputRow, masCreatedDateCol).Value >= NEWJOBDATE Then
				'''''''''''''''''''Find job in Nate's Sheet'''''''''''''''



				'''''''''''''''''''Print out the Output'''''''''''''''''''
				With printSheet
					''''''From Master Report'''''
					.Cells(printRow, customerCol).Value     = masterInput.Cells(inputRow, masCustomerCol).Value
					.Cells(printRow, jobCol).Value          = masterInput.Cells(inputRow, masJobCol).Value
					.Cells(printRow, kWCol).Value           = masterInput.Cells(inputRow, masKWCol).Value
					.Cells(printRow, statusCol).Value       = masterInput.Cells(inputRow, masStatusCol).Value
					.Cells(printRow, subStatusCol).Value    = masterInput.Cells(inputRow, masSubStatusCol).Value
					.Cells(printRow, createdDateCol).Value  = masterInput.Cells(inputRow, masCreatedDateCol).Value				
					.Cells(printRow, repEmailCol).Value     = masterInput.Cells(inputRow, masRepEmailCol).Value
					.Cells(printRow, finalContratCol).Value = masterInput.Cells(inputRow, masFinalCol).Value
					.Cells(printRow, stateCol).Value        = masterInput.Cells(inputRow, masStateCol).Value
					printRow = printRow + 1
				End With
			End If

		Next inputRow


End Sub
