function secondPayment(ByRef currentJob As cJobData, ByRef currentRep as cRepData, ByVal WorkBookName As String) As cJobData

'Happens after final contract is signed
dim second_payment_rate as currency
	second_payment_rate = 50
dim number_of_kW as integer
dim second_payment_total as currency

'Need to find the total number of kW that qualify for this week's second payment.
second_payment_total = currentJob.kW * second_payment_rate

'determine if job is at second payment status and that it hasn't been cancelled'
If currentJob.isFinalContract AND NOT currentJob.isCancelled Then
	currentJob.ThisWeekSecondPayment = second_payment_total
	printSecond currentJob, currentRep, workBookName
End If

	Set secondPayment = currentJob

End function

Sub printSecond(ByRef currentJob As cJobData,  ByRef currentRep as cRepData, ByVal WorkBookName As String)
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