Sub getCountInfo(ByVal workBookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim customerCol, jobCol, kWCol, createdDateCol, _
        paymentAmountCol, paymentDateCol, _
        secondPaymentCol, secondDateCol, _
        finalPaymentCol, finalDateCol As Integer

        customerCol      = 3
        jobCol           = 4
        kWCol            = 5
        createdDateCol   = 6
        paymentAmountCol = 7
        paymentDateCol   = 8
        secondPaymentCol = 9
        secondDateCol    = 10
        finalPaymentCol  = 12
        finalDateCol     = 13
        
''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim firstPaymentSheet, secondPaymentSheet, InstalledSheet As Worksheet
        Set firstPaymentSheet = NatesEvolution.Worksheets("1st_Payments_Pending")
        Set secondPaymentSheet = NatesEvolution.Worksheets("2nd_Payments_Pending")
        Set InstalledSheet = NatesEvolution.Worksheets("Installed")
''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
    Dim printRow As Integer
        printRow = 2
'''''''''''''''''''''''''''''job Object''''''''''''''''''''''
    Dim rep As cRepData
    Dim job As cJobData
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''GET PAYMENT INFORMATION'''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Find jobs in the "1st_Payments_Pending" and'
    '"2nd_Payments_Pending" Tabs and update the jobData info'

    For jobIndex = 0 To UBound(payroll_main.jobData)
            Set job = payroll_main.jobData(jobIndex)
            On Error Resume Next
            Set rep = payroll_main.repData(job.repEmail)	

			If job.firstPaymentAmount = 0 and job.repEmail = rep.Email and job.Status <> "Cancelled" then
				if job.CreatedDate > rep.FirstJobDate + 60 and job.IsFinalContract = True then
					rep.SalesThisWeek = rep.SalesThisWeek + 1
				elseif job.CreatedDate < rep.FirstJobDate + 60 and job.IsSurveyComplete = True then
					rep.SalesThisWeek = rep.SalesThisWeek + 1
				End if
			End if

		 'Reset the job in the array'
        Set payroll_main.repData.Item(repIndex) = rep
    Next

End Sub