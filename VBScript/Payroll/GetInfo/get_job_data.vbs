Function getJobData(ByVal workBookName As String) As cJobData()

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
        repEmailCol        = 7
        isFinalContractCol = 8
        stateCol           = 9
        
    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

    ''''''''''''''''''''''''''Create Nate's Evolution'''''''''''''
    	createNatesEvo(workBookName)
    	
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim jobDataSheet, printSheet As Worksheet
        Set jobDataSheet = NatesEvolution.Worksheets("Nate's Evolution")
        Set printSheet = NatesEvolution.Worksheets("Debug")

    ''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
    Dim inputRow, dataRow As Integer
        dataRow = 0
    '''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
    Dim jobData() As cJobData
    Dim currentJob  As cJobData

    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim jobDataSize As Long
    	jobDataSize = jobDataSheet.Cells(1,1).End(xlDown).Row - 2
    	ReDim jobData(jobDataSize)

   	'''''''''''Date constant to consider for new jobs''''''''''''''
   	Const NEWJOBDATE = #11/30/2014#

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''GET AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      For inputRow = 2 To jobDataSize + 2
            With jobDataSheet
                
            	'determine if the job is in the new pay structure'
                If  .Cells(inputRow, createdDateCol).value >= NEWJOBDATE Then

                	'Get data from sheet and pass it to new Data object'
	                Set currentJob = New cJobData
	                    currentJob.Customer    = .Cells(inputRow, customerCol).value
	                    currentJob.JobID       = .Cells(inputRow, jobCol).value
	                    currentJob.kW          = .Cells(inputRow, kWCol).value
	                    currentJob.Status      = .Cells(inputRow, statusCol).value
	                    currentJob.SubStatus   = .Cells(inputRow, subStatusCol).value
	                    currentJob.CreatedDate = .Cells(inputRow, createdDateCol).value
	                    currentJob.RepEmail    = .Cells(inputRow, repEmailCol).value
                        currentJob.States       = .Cells(inputRow, stateCol).value
                        currentJob.setIsFinalContract(.Cells(inputRow, isFinalContractCol).value)
	                    currentJob.setIsInstall
	                    currentJob.setIsCancelled
                        currentJob.setDaysSinceCreated
                        currentJob.setIsSurveyComplete

	                ''''''''''Add currentJob to the jobData Array'''''''''''''
                    IF currentJob.IsInstall Or currentJob.isFinalContract Or currentJob.IsSurveyComplete Then
	                   Set jobData(dataRow) = currentJob
                       dataRow = dataRow + 1
                    End If

                End If
            End With

        Next inputRow

        getJobData = jobData
        
        


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''PRINT VALUES''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function




