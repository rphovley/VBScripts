Sub getFromMaster()

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
	    repEmailCol        = 9
	    isDocSignedCol     = 7
	    isFinalContractCol = 8

	''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
		Dim workBookName As String
		workBookName = inputBox("What is the master report's name?") & ".xlsx"
		Dim NatesEvolution As Workbook
		Set NatesEvolution = Workbooks(workBookName)

	''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
		Dim inputDataSheet, printSheet As Worksheet
		Set inputDataSheet = NatesEvolution.Worksheets("Current Data")
		Set printSheet = NatesEvolution.Worksheets("TestPrint")

	''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
		Dim inputRow, printRow As Integer
		printRow = 2

	'''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
		Dim inputData() As cJobData
		Dim currentRep  As cJobData
		Dim printRep    As cJobData

	'''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
		Dim inputDataSize As Long
		inputDataSize = inputDataSheet.UsedRange.Rows.Count - 1
		ReDim inputData(inputDataSize)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''GET AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

		For inputRow = 2 To inputDataSize
			With inputDataSheet
				'Get data from sheet and pass it to new Data object'
				Set currentRep = New cJobData
					currentRep.Customer        = .Cells(inputRow, customerCol).Value
					currentRep.JobID           = .Cells(inputRow, jobCol).Value
					currentRep.kW              = .Cells(inputRow, kWCol).Value
					currentRep.Status          = .Cells(inputRow, statusCol).Value 
					currentRep.SubStatus       = .Cells(inputRow, subStatusCol).Value
					currentRep.CreatedDate     = .Cells(inputRow, createdDateCol).Value
					currentRep.RepEmail        = .Cells(inputRow, repEmailCol).Value
					currentRep.setIsDocSigned(.Cells(inputRow, isDocSignedCol).Value)
					currentRep.setIsFInalContract(.Cells(inputRow, isFinalContractCol).Value)
					currentRep.setIsInstall
					currentRep.setIsCancelled
			End With

			Set inputData(inputRow - 2) = currentRep
		Next inputRow

		For i = 0 To inputDataSize

			With printSheet
					.Cells(i + 2, 1).Value = inputData(i).Customer
					.Cells(i + 2, 2).Value = inputData(i).JobID  
					.Cells(i + 2, 3).Value = inputData(i).kW 
					.Cells(i + 2, 4).Value = inputData(i).CreatedDate 
					.Cells(i + 2, 5).Value = inputData(i).Status
					.Cells(i + 2, 6).Value = inputData(i).SubStatus 
					.Cells(i + 2, 7).Value = inputData(i).RepEmail
					.Cells(i + 2, 8).Value = inputData(i).isDocSigned
					.Cells(i + 2, 9).Value = inputData(i).isFinalContract
					.Cells(i + 2, 10).Value = inputData(i).isInstall
					.Cells(i + 2, 11).Value = inputData(i).isCancelled

			End With
		Next i


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''PRINT VALUES TESt'''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub