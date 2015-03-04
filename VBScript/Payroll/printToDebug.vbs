Sub printToDebug(ByRef jobData() As cJobData, ByVal workbookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
        createdDateCol, repEmailCol, isDocSignedCol, isFinalContractCol, _
        isInstallCol, isCancelledCol, firstPayCol, firstPayDateCol, _
        secondPayCol, secondPayDateCol, whatWasPaidCol As Integer

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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''PRINT JOBS TO DEBUG SHEET'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	For Each printJob In jobData
            
            With printSheet
            
                    .Cells(printRow, customerCol).value         = printJob.Customer
                    .Cells(printRow, jobCol).value              = printJob.JobID
                    .Cells(printRow, kWCol).value               = printJob.kW
                    .Cells(printRow, createdDateCol).value      = printJob.CreatedDate
                    .Cells(printRow, statusCol).value           = printJob.Status
                    .Cells(printRow, subStatusCol).value        = printJob.SubStatus
                    .Cells(printRow, repEmailCol).value         = printJob.RepEmail
                    .Cells(printRow, isDocSignedCol).value      = printJob.IsDocSigned
                    .Cells(printRow, isFinalContractCol).value  = printJob.IsFinalContract
                    .Cells(printRow, isInstallCol).value        = printJob.IsInstall
                    .Cells(printRow, isCancelledCol).value      = printJob.IsCancelled
                    .Cells(printRow, firstPayCol).value         = printJob.FirstPaymentAmount
                    .Cells(printRow, firstPayDateCol).value     = printJob.FirstPaymentDate
                    .Cells(printRow, secondPayCol).value        = printJob.SecondPaymentAmount
                    .Cells(printRow, secondPayDateCol).value    = printJob.SecondPaymentDate
                    .Cells(printRow, whatWasPaidCol).value      = printJob.WhatWasPaid

            End With
            
            printRow = printRow + 1
        Next

End Sub
