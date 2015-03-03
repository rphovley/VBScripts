Function getPaymentInfo(ByRef jobData() As cJobData, ByVal workbookName As String) As cJobData()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim customerCol, jobCol, kWCol, createdDateCol, _
    	paymentAmountCol, paymentDateCol As Integer

        customerCol      = 3
        jobCol           = 4
        kWCol            = 5
        createdDateCol   = 6
        paymentAmountCol = 7
        paymentDateCol   = 8

        
''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim firstPaymentSheet, secondPaymentSheet, InstalledSheet As Worksheet
        Set firstPaymentSheet  = NatesEvolution.Worksheets("1st_Payments_Pending")
        Set secondPaymentSheet = NatesEvolution.Worksheets("2nd_Payments_Pending")
        Set InstalledSheet     = NatesEvolution.Worksheets("Installed")
''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
    Dim printRow As Integer
        printRow = 2
'''''''''''''''''''''''''''''job Object''''''''''''''''''''''
	Dim job As cJobData	

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''GET PAYMENT INFORMATION'''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	'Find jobs in the "1st_Payments_Pending" and'
	'"2nd_Payments_Pending" Tabs and update the jobData info'
	For jobIndex = 0 To UBound(jobData)
		set job = jobData(jobIndex)

		 With InstalledSheet

	     	'Loop through the first Payments Sheet'
	     	For jobRow = 2 To .Cells(1,1).End(xlDown).Row

	     		'If this is true, this job is in the 2nd_Payments_Pending Tab'
	     		If .Cells(jobRow, jobCol).Value = job.JobID Then
	     			'update jobData with new information'
	     			job.IsPaidInFull = True
	     		End If


	     	Next jobRow
	     End With 

	     'Do not check for job in first and second payment tab if it is already paid in full'
	     If NOT job.IsPaidInFull Then
	         With firstPaymentSheet

	         	'Loop through the first Payments Sheet'
	         	For jobRow = 2 To .Cells(1,1).End(xlDown).Row

	         		'If this is true, this job is in the 1st_Payments_Pending Tab'
	         		If .Cells(jobRow, jobCol).Value = job.JobID Then
	         			'update jobData with new information'
	         			job.FirstPaymentAmount  = .Cells(jobRow, paymentAmountCol).Value
	         			job.FirstPaymentDate    = .Cells(jobRow, paymentDateCol).Value
	         		End If


	         	Next jobRow
	         End With  
	            
	         With secondPaymentSheet

	         	'Loop through the first Payments Sheet'
	         	For jobRow = 2 To .Cells(1,1).End(xlDown).Row

	         		'If this is true, this job is in the 2nd_Payments_Pending Tab'
	         		If .Cells(jobRow, jobCol).Value = job.JobID Then
	         			'update jobData with new information'
	         			job.SecondPaymentAmount = .Cells(jobRow, paymentAmountCol).Value
	         			job.SecondPaymentDate   = .Cells(jobRow, paymentDateCol).Value
	         		End If


	         	Next jobRow
	         End With   
	     End If
          
         'set the value of what was paid'
         job.setWhatWasPaid
    Next



End Function