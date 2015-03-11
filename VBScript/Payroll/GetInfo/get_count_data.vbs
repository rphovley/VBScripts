Function getCountInfo(ByRef jobData() As cJobData, ByRef repData() As cRepData, ByVal workBookName As String) As cRepData()

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
    Dim job As cJobData
    Dim newJobData() As cJobData
    ReDim newJobData(UBound(jobData))

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''GET PAYMENT INFORMATION'''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Find jobs in the "1st_Payments_Pending" and'
    '"2nd_Payments_Pending" Tabs and update the jobData info'
    For repIndex = 0 To UBound(repData)
        Set rep = repData.Item(repIndex)
         Set rep = findRep(repData, rep.Email)

		 For jobIndex = 0 To UBound(jobData)
		 
			If job.firstPaymentAmount = 0 and job.repEmail = rep then
				SalesThisWeek = SalesThisWeek + 1
			End if

		 Next
		 'Reset the job in the array'
        Set repData.Item(repIndex) = rep
    Next


    getPaymentInfo = jobData


End Function

'returns the rep object associated with the job'
Function findRep( ByRef repData As Collection, ByVal repEmail As String) As cRepData
        
    For Each rep In repData
        If rep.Email = repEmail Then
            Set findRep = rep
        End IF
    Next

End Function
