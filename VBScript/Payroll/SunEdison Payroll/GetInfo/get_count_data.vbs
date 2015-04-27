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
    Dim rep     As cRepData
    Dim job     As cJobData
    Dim weather As cWeatherData
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''GET PAYMENT INFORMATION'''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Find jobs in the "1st_Payments_Pending" and'
    '"2nd_Payments_Pending" Tabs and update the jobData info'

    For jobIndex = 0 To UBound(payroll_main.jobData)
            Set rep     = Nothing
            Set job     = Nothing
            Set weather = Nothing
            Set job = payroll_main.jobData(jobIndex)
            On Error Resume Next
            Set rep     = payroll_main.repData(job.repEmail)
            Set weather = payroll_main.weatherData(job.repEmail)	

    		If Not weather is Nothing OR DateDiff("d", rep.FirstJobDate, job.CreatedDate) <= 60 Then
                If job.WhatWasPaid = 0 AND NOT rep is Nothing AND NOT job.IsCancelled _ 
                    AND NOT rep.IsBlackList AND job.IsSurveyComplete Then
    					rep.SalesThisWeek = rep.SalesThisWeek + 1
    			End if
            Else
                If job.IsDocSigned And job.WhatWasPaid = 0 And Not rep Is Nothing And Not job.IsCancelled _
                    And Not rep.IsBlackList And job.isSurveyComplete Then
                    rep.SalesThisWeek = rep.SalesThisWeek + 1
                End If
            End If 
		 'Reset the job in the array'
        Set payroll_main.repData.Item(repIndex) = rep
    Next

End Sub