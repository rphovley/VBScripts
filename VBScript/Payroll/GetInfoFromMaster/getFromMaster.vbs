Sub getFromMaster()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
	Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
	    createdDateCol, repEmailCol As Integer

	    customerCol    = 1   
	    jobCol         = 2
	    kWCol          = 3
	    statusCol      = 4
	    subStatusCol   = 5
	    createdDateCol = 7
	    repEmailCol    = 17

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
		Dim inputData() As jobDataClass
		Dim currentRep  As jobDataClass
		Dim printRep    As jobDataClass

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
				Set currentRep = New jobDataClass
					currentRep.Customer    = .Cells(inputRow, customerCol).Value
					currentRep.JobID       = .Cells(inputRow, jobCol).Value
					currentRep.kW          = .Cells(inputRow, kWCol).Value
					currentRep.Status      = .Cells(inputRow, statusCol).Value 
					currentRep.SubStatus   = .Cells(inputRow, subStatusCol).Value
					currentRep.CreatedDate = .Cells(inputRow, createdDateCol).Value
					currentRep.RepEmail    = .Cells(inputRow, repEmailCol).Value

			End With

			Set inputData(inputRow - 2) = currentRep
		Next inputRow

		For i = 0 To inputDataSize

			With printSheet
					.Cells(inputRow, 1).Value = inputData(i).Customer
					.Cells(inputRow, 2).Value = inputData(i).JobID  
					.Cells(inputRow, 3).Value = inputData(i).kW 
					.Cells(inputRow, 4).Value = inputData(i).CreatedDate 
					.Cells(inputRow, 5).Value = inputData(i).Status
					.Cells(inputRow, 6).Value = inputData(i).SubStatus 
					.Cells(inputRow, 7).Value = inputData(i).RepEmail
			End With
		Next i


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''PRINT VALUES TESt''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

End Sub