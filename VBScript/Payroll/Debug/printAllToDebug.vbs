Sub printAllToDebug(ByVal workBookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
        createdDateCol, repEmailCol, isDocSignedCol, isFinalContractCol, _
        isInstallCol, isCancelledCol, firstPayCol, firstPayDateCol,  _
        secondPayCol, secondPayDateCol, finalPayCol,  finalDateCol, _
        whatWasPaidCol, thisWeeksFirstCol, thisWeeksSecondCol, _
        thisWeeksFinalCol, thisCancellationCol, _
        daysSinceCol, repNameCol, repScaleCol, _
        repBlackCol, repInactiveCol, repNewCol, _
        repSliderCol, repSliderDateCol, repMarketEndCol, _
        repMarketStartCol, repIsMarketCol, repMarketRateCol, _
        repFirstJobCol, repSalesWeekCol, repInstallPool, isSurveyComplete As Integer

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
        finalPayCol        = 16
        finalDateCol       = 17
        whatWasPaidCol     = 18
        thisWeeksFirstCol  = 19
        thisWeeksSecondCol = 20
        thisWeeksFinalCol  = 21
        thisCancellationCol= 22
        daysSinceCol       = 23
        repNameCol         = 24
        repScaleCol        = 25
        repBlackCol        = 26
        repInactiveCol     = 27
        repNewCol          = 28
        repSliderCol       = 29
        repSliderDateCol   = 30
        repIsMarketCol     = 31
        repMarketStartCol  = 32
        repMarketEndCol    = 33 
        repMarketRateCol   = 34
        repFirstJobCol     = 35
        repSalesWeekCol    = 36
        repInstallPool     = 37
        isSurveyComplete   = 38
        
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

        ''''''object''''
        Dim rep As cRepData

    Const EMPTYDATE = #12:00:00 AM#
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''PRINT JOBS TO DEBUG SHEET'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For Each printJob In payroll_main.jobData
            On Error Resume Next
            Set rep = payroll_main.repData(printJob.repEmail)
            With printSheet
            
                    .Cells(printRow, customerCol).value        = printJob.Customer
                    .Cells(printRow, jobCol).value             = printJob.JobID
                    .Cells(printRow, kWCol).value              = printJob.kW
                    .Cells(printRow, createdDateCol).value     = printJob.CreatedDate
                    .Cells(printRow, statusCol).value          = printJob.Status
                    .Cells(printRow, subStatusCol).value       = printJob.SubStatus
                    .Cells(printRow, repEmailCol).value        = printJob.RepEmail
                    .Cells(printRow, isDocSignedCol).value     = printJob.IsDocSigned
                    .Cells(printRow, isFinalContractCol).value = printJob.IsFinalContract
                    .Cells(printRow, isInstallCol).value       = printJob.IsInstall
                    .Cells(printRow, isCancelledCol).value     = printJob.IsCancelled
                    .Cells(printRow, isSurveyComplete).value   = printJob.IsSurveyComplete
                    
                    'Leave the cell empty if it equals Empty Date'
                    If printJob.FirstPaymentDate <> EMPTYDATE Then
                        .Cells(printRow, firstPayCol).value     = printJob.FirstPaymentAmount
                        .Cells(printRow, firstPayDateCol).value = printJob.FirstPaymentDate
                    End If

                    'Leave the cell empty if it equals Empty Date'
                    If printJob.SecondPaymentDate <> EMPTYDATE Then
                        .Cells(printRow, secondPayCol).value     = printJob.SecondPaymentAmount
                        .Cells(printRow, secondPayDateCol).value = printJob.SecondPaymentDate
                    End If

                    'Leave the cell empty if it equals Empty Date'
                    If printJob.FinalPaymentDate <> EMPTYDATE Then
                        .Cells(printRow, finalPayCol).value     = printJob.FinalPaymentAmount
                        .Cells(printRow, finalDateCol).value    = printJob.FinalPaymentDate
                    End If

                    'Leave the cell empty if it equals 0'
                    If printJob.ThisWeekFirstPayment <> 0 Then
                        .Cells(printRow, thisWeeksFirstCol).value = printJob.ThisWeekFirstPayment
                    End If

                    'Leave the cell empty if it equals 0'
                    If printJob.ThisWeekSecondPayment <> 0 Then
                        .Cells(printRow, thisWeeksSecondCol).value = printJob.ThisWeekSecondPayment
                    End If        

                    'Leave the cell empty if it equals 0'
                    If printJob.ThisWeekFinalPayment <> 0 Then
                        .Cells(printRow, thisWeeksFinalCol).value = printJob.ThisWeekFinalPayment
                    End If

                    'Leave the cell empty if it equals 0'
                    If printJob.ThisWeekCancelled <> 0 Then
                        .Cells(printRow, thisCancellationCol).value     = printJob.ThisWeekCancelled
                    End If

                    
                    .Cells(printRow, whatWasPaidCol).value = printJob.WhatWasPaid

                    .Cells(printRow, daysSinceCol).value = printJob.DaysSinceCreated

                If rep.Email = printJob.RepEmail Then
                    .Cells(printRow, repNameCol).value       = rep.Name
                    .Cells(printRow, repScaleCol).value      = rep.PayScaleID
                    .Cells(printRow, repBlackCol).value      = rep.IsBLackList
                    .Cells(printRow, repInactiveCol).value   = rep.IsInactive
                    .Cells(printRow, repFirstJobCol).Value   = rep.FirstJobDate
                    .Cells(printRow, repSalesWeekCol).Value  = rep.SalesThisWeek
                    .Cells(printRow, repInstallPool).Value   = rep.InstallPool
                    If rep.IsSlider Then
                        .Cells(printRow, repSliderCol).value     = rep.IsSlider
                        .Cells(printRow, repSliderDateCol).value = rep.StartSliderDate
                    End If

                    'check if the rep is a market rep'
                    If rep.IsMarketing Then
                        .Cells(printRow, repIsMarketCol).Value    = rep.IsMarketing
                        .Cells(printRow, repMarketStartCol).Value = rep.MarkStartDate
                        .Cells(printRow, repMarketEndCol).Value   = rep.MarkEndDate
                        .Cells(printRow, repMarketRateCol).Value  = rep.MarketingRate
                    End If
                End If

            End With
            printRow = printRow + 1
        Next

End Sub


