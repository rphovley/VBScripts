Function getPaymentInfo(ByRef jobData() As cJobData, ByVal workBookName As String) As cJobData()

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
    For jobIndex = 0 To UBound(jobData)
        Set job = jobData(jobIndex)

         With InstalledSheet

            'Loop through the first Payments Sheet'
            For jobRow = 2 To .Cells(1, 1).End(xlDown).Row

                'If this is true, this job is in the 2nd_Payments_Pending Tab'
                If .Cells(jobRow, jobCol).value = job.JobID Then
                    'update jobData with new information'
                    job.IsPaidInFull = True
                    job.FinalPaymentAmount  = .Cells(jobRow, finalPaymentCol).value
                    job.FinalPaymentDate    = .Cells(jobRow, finalDateCol).value
                    job.FirstPaymentAmount  = .Cells(jobRow, paymentAmountCol).value
                    job.FirstPaymentDate    = .Cells(jobRow, paymentDateCol).value
                    job.SecondPaymentAmount = .Cells(jobRow, secondPaymentCol).value
                    job.SecondPaymentDate   = .Cells(jobRow, secondDateCol).value

                End If


            Next jobRow

         End With

         'Do not check for job in first and second payment tab if it is already paid in full'
         If Not job.IsPaidInFull Then
             With firstPaymentSheet

                'Loop through the first Payments Sheet'
                For jobRow = 2 To .Cells(1, 1).End(xlDown).Row

                    'If this is true, this job is in the 1st_Payments_Pending Tab'
                    If .Cells(jobRow, jobCol).value = job.JobID Then
                        'update jobData with new information'
                        job.FirstPaymentAmount = .Cells(jobRow, paymentAmountCol).value
                        job.FirstPaymentDate   = .Cells(jobRow, paymentDateCol).value
                    End If


                Next jobRow
             End With
                
             With secondPaymentSheet

                'Loop through the first Payments Sheet'
                For jobRow = 2 To .Cells(1, 1).End(xlDown).Row

                    'If this is true, this job is in the 2nd_Payments_Pending Tab'
                    If .Cells(jobRow, jobCol).value = job.JobID Then
                        'update jobData with new information'
                        job.SecondPaymentAmount = .Cells(jobRow, paymentAmountCol).value
                        job.SecondPaymentDate   = .Cells(jobRow, paymentDateCol).value
                    End If


                Next jobRow
             End With
         End If


         Set rep = findRep(repData, job.repEmail)

        'set the value of what was paid'
        job.setWhatWasPaid

        'Reset the job in the array'
        Set jobData(jobIndex) = job
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
