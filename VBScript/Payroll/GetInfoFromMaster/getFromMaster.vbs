Function getJobData(ByVal workBookName As String) As cJobData()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
        createdDateCol, repEmailCol, isDocSignedCol, isFinalContractCol As Integer

        customerCol = 1
        jobCol = 2
        kWCol = 3
        statusCol = 4
        subStatusCol = 5
        createdDateCol = 6
        repEmailCol = 9
        isDocSignedCol = 7
        isFinalContractCol = 8
        
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
    Dim inputRow, printRow As Integer
        printRow = 2

    '''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
    Dim jobData() As cJobData
    Dim currentRep  As cJobData

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
	                Set currentRep = New cJobData
	                    currentRep.Customer    = .Cells(inputRow, customerCol).value
	                    currentRep.JobID       = .Cells(inputRow, jobCol).value
	                    currentRep.kW          = .Cells(inputRow, kWCol).value
	                    currentRep.Status      = .Cells(inputRow, statusCol).value
	                    currentRep.SubStatus   = .Cells(inputRow, subStatusCol).value
	                    currentRep.CreatedDate = .Cells(inputRow, createdDateCol).value
	                    currentRep.RepEmail    = .Cells(inputRow, repEmailCol).value
	                    currentRep.setIsDocSigned (.Cells(inputRow, isDocSignedCol).value)
	                    currentRep.setIsFinalContract (.Cells(inputRow, isFinalContractCol).value)
	                    currentRep.setIsInstall
	                    currentRep.setIsCancelled

	                ''''''''''Add currentRep to the jobData Array'''''''''''''
	                Set jobData(inputRow - 2) = currentRep

                End If
            End With

            
        Next inputRow

        getJobData = jobData
        
        


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''PRINT VALUES''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function




