function cancellations(ByRef currentJob As cJobData, ByRef currentRep as cRepData, ByVal WorkBookName As String) As cJobData
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''COLUMNS'''''
Dim jobCol, emailCol, kWCol, paidCol, clawedCol, dateCol As Integer
	
	jobCol    = 1
	emailCol  = 2
	kWCol     = 3
	paidCol   = 4
	clawedCol = 5
	dateCol   = 6

Dim row As Integer
	row = 2 

Dim isAlreadyCancelled As Boolean

'Check if job is now cancelled'
	If currentJob.IsCancelled Then

	With WorkBook(WorkBookName).WorkSheet("Cancelled")
		Do Until IsEmpty(.Cells(row, 1))
			'If it is cancelled, check if is already on the list of jobs cancelled check to see if there are installs to take from'
			If .Cells(row, jobCol) = currentJob.JobID Then
				If .Cells(row, paidCol).Value - .Cells(row, clawedCol).Value = 0 Then


					isAlreadyCancelled = True
					Exit Do
				End If
			End If
			row = row + 1
		Loop

		'if the job is cancelled and is not already clawed back add an entry to cancellation sheet'
		If Not isAlreadyCancelled Then

		End If

	End With

'If we claw back the job, remove it from any of the payment sheets so the "balance" for the job will now be 0'

'Print out to debug sheet what was actually clawed back and what was not'

	End If



'determine if job is at second payment status and that it hasn't been cancelled'
If currentJob.isFinalContract AND NOT currentJob.isCancelled And Not currentRep.IsBlackList Then
	
	If currentRep.IsMarketing Then
		second_payment_rate = 25
	Else
		second_payment_rate = 50
	End If

	second_payment_total = currentJob.kW * second_payment_rate

	currentJob.ThisWeekSecondPayment = second_payment_total
	printSecond currentJob, currentRep, workBookName
End If

	Set secondPayment = currentJob

End function

Sub printCancellation(ByRef currentJob As cJobData,  ByRef currentRep as cRepData, ByVal WorkBookName As String)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim repCol, repIDCol, customerCol, jobCol, kWCol, createdDateCol, _
        paymentAmountCol, paymentDateCol As Integer

        repCol           = 1
        repIDCol         = 2
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
	
	'worksheets'
	Dim secondPayment As Worksheet
		Set secondPayment = NatesEvolution.Worksheets("2nd_Payments_Pending")
	'Row'
	Dim printRow As Integer
		printRow = secondPayment.Cells(1,1).End(xlDown).Row + 1

	With secondPayment
		.Cells(printRow, repCol)           = currentRep.Name
		.Cells(printRow, repIDCol)         = currentRep.ID
		.Cells(printRow, customerCol)      = currentJob.Customer
		.Cells(printRow, jobCol)           = currentJob.JobID
		.Cells(printRow, kWCol)            = currentJob.kW
		.Cells(printRow, createdDateCol)   = currentJob.CreatedDate
		.Cells(printRow, paymentAmountCol) = currentJob.ThisWeekSecondPayment
		.Cells(printRow, paymentDateCol)   = Date
	End With

End Sub