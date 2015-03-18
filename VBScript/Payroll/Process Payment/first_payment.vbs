Function firstPayment(ByRef currentRep As cRepData, ByRef currentJob As cJobData, ByVal WorkBookName As String) As cJobData

''Occurs when site survey complete or 2 weeks after job created date
Dim first_payment_total As Currency

Dim new_one_two As Currency
    new_one_two = 100
Dim new_three_five As Currency
    new_three_five = 200
Dim new_six_plus As Currency
    new_six_plus = 300
    
''Occurs when final contract is signed
Dim old_one_two As Currency
    old_one_two = 250
Dim old_three_five As Currency
    old_three_five = 350
Dim old_six_plus As Currency
    old_six_plus = 450
Dim first_payment As Currency

'Needs to first count how many accounts qualify for this week's  first payment

 ''''''''''''''''''''''''''''''Calculates the first payment''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FOR ACCOUNTS LESS THAN 60 DAYS''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
if currentJob.IsSurveyComplete AND NOT currentJob.isCancelled And Not currentRep.IsBlackList then

	If DateDiff("d", rep.FirstJobDate, job.CreatedDate) <= 60 Then
		If SalesThisWeek <= 2 And SalesThisWeek > 0 Then
			first_payment_total = new_one_two
		ElseIf SalesThisWeek > 2 And SalesThisWeek <= 5 Then
			first_payment_total = new_three_five
		ElseIf SalesThisWeek > 5 Then
			first_payment_total = new_six_plus
		End If
	Else

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''FOR ACCOUNTS GREATER THAN 60 DAYS''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		If SalesThisWeek <= 2 And SalesThisWeek > 0 Then
			first_payment_total = old_one_two
		ElseIf SalesThisWeek > 2 And SalesThisWeek <= 5 Then
			first_payment_total = old_three_five
		ElseIf SalesThisWeek > 5 Then
			first_payment_total = old_six_plus
		End If
	End If
		currentJob.ThisWeekFirstPayment = first_payment_total
		printFirst currentJob, currentRep, workBookName
End If	
	
    Set firstPayment = currentJob
    
End Function
        
        
        
Sub printFirst(ByRef currentRep As cRepData)
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''''''Columns'''''''''''''''''''''''''
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
        Set NatesEvolution = Workbooks(WorkBookName)

    ''''''''''''''''''''''''''Create Nate's Evolution'''''''''''''
        createNatesEvo (WorkBookName)
        
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim firstPayment As Worksheet
        Set firstPayment = NatesEvolution.Worksheets("1st_Payments_Pending")

    ''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
    Dim printRow As Integer
        printRow = firstPayment.cells(1,1).End(xlDown).Row +1

	With firstPayment
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