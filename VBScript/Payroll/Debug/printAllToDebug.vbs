Sub printAllToDebug(ByRef jobData() As cJobData, ByRef repData As Collection, ByVal workBookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
        createdDateCol, repEmailCol, isDocSignedCol, isFinalContractCol, _
        isInstallCol, isCancelledCol, firstPayCol, firstPayDateCol, _
        secondPayCol, secondPayDateCol, whatWasPaidCol, daysSinceCol, _
        repNameCol, repScaleCol, repBlackCol, repInactiveCol, repNewCol, _
        repSliderCol, repSliderDateCol As Integer

        customerCol        = 1
        jobCol             = 2
        kWCol              = 3
        createdDateCol     = 4
        statusCol          = 5
        subStatusCol       = 6
        repEmailCol        = 7
        isDocSignedCol     = 8
        isFinalContractCol = 9
        isInstallCol       = 10
        isCancelledCol     = 11
        firstPayCol        = 12
        firstPayDateCol    = 13
        secondPayCol       = 14
        secondPayDateCol   = 15
        whatWasPaidCol     = 16
        daysSinceCol       = 17
        repNameCol         = 18
        repScaleCol        = 19
        repBlackCol        = 20
        repInactiveCol     = 21
        repNewCol          = 22
        repSliderCol       = 23
        repSliderDateCol   = 24
        
''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim printSheet As Worksheet
        Set printSheet = NatesEvolution.Worksheets("Debug")

''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
    Dim printRow As Integer
        printRow = 2

    Const EMPTYDATE = #12:00:00 AM#
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''PRINT JOBS TO DEBUG SHEET'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For Each printJob In jobData
            
            With printSheet
            
                    .Cells(printRow, customerCol).value = printJob.Customer
                    .Cells(printRow, jobCol).value = printJob.JobID
                    .Cells(printRow, kWCol).value = printJob.kW
                    .Cells(printRow, createdDateCol).value = printJob.CreatedDate
                    .Cells(printRow, statusCol).value = printJob.Status
                    .Cells(printRow, subStatusCol).value = printJob.SubStatus
                    .Cells(printRow, repEmailCol).value = printJob.RepEmail
                    .Cells(printRow, isDocSignedCol).value = printJob.IsDocSigned
                    .Cells(printRow, isFinalContractCol).value = printJob.IsFinalContract
                    .Cells(printRow, isInstallCol).value = printJob.IsInstall
                    .Cells(printRow, isCancelledCol).value = printJob.IsCancelled
                    
                    'Leave the cell empty if it equals Empty Date'
                    If printJob.FirstPaymentDate <> EMPTYDATE Then
                        .Cells(printRow, firstPayCol).value = printJob.FirstPaymentAmount
                        .Cells(printRow, firstPayDateCol).value = printJob.FirstPaymentDate
                    End If

                    'Leave the cell empty if it equals Empty Date'
                    If printJob.SecondPaymentDate <> EMPTYDATE Then
                        .Cells(printRow, secondPayCol).value = printJob.SecondPaymentAmount
                        .Cells(printRow, secondPayDateCol).value = printJob.SecondPaymentDate
                    End If
                    .Cells(printRow, whatWasPaidCol).value = printJob.WhatWasPaid

                    .Cells(printRow, daysSinceCol).value = printJob.DaysSinceCreated

            
            
            'Get rep info for the job'
            For Each rep In repData
                If rep.Email = printJob.RepEmail Then
                    .Cells(printRow, repNameCol).value       = rep.Name
                    .Cells(printRow, repScaleCol).value      = rep.PayScaleID
                    .Cells(printRow, repBlackCol).value      = rep.IsBLackList
                    .Cells(printRow, repInactiveCol).value   = rep.IsInactive
                    .Cells(printRow, repNewCol).value        = rep.IsNewRep
                    .Cells(printRow, repSliderCol).value     = rep.IsSlider
                    .Cells(printRow, repSliderDateCol).value = rep.StartSliderDate
                End If
            Next

            End With
            printRow = printRow + 1
        Next

End Sub


