Sub printToDebug(ByRef jobData() As cJobData, ByVal workbookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

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
                    .Cells(printRow, 1).value  = printJob.Customer
                    .Cells(printRow, 2).value  = printJob.JobID
                    .Cells(printRow, 3).value  = printJob.kW
                    .Cells(printRow, 4).value  = printJob.CreatedDate
                    .Cells(printRow, 5).value  = printJob.Status
                    .Cells(printRow, 6).value  = printJob.SubStatus
                    .Cells(printRow, 7).value  = printJob.RepEmail
                    .Cells(printRow, 8).value  = printJob.IsDocSigned
                    .Cells(printRow, 9).value  = printJob.IsFinalContract
                    .Cells(printRow, 10).value = printJob.IsInstall
                    .Cells(printRow, 11).value = printJob.IsCancelled

            End With
            
            printRow = printRow + 1
        Next

End Sub
