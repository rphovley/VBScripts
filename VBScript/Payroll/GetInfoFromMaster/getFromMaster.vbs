Sub getFromMaster()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
        createdDateCol, repEmailCol, isDocSignedCol, isFinalContractCol As Integer

        customerCol = 1
        jobCol = 2
        kWCol = 3
        statusCol = 4
        subStatusCol = 5
        createdDateCol = 6
        repEmailCol = 9
        isDocSignedCol = 7
        isFinalContractCol = 8
        
    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        Dim workBookName As String
        Dim testName As String
        testName = "VBA Triforce (Ezra)"
        workBookName = testName & ".xlsm"
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
        Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

    ''''''''''''''''''''''''''Create Nate's Evolution'''''''''''''
    	createNatesEvo(workBookName)
    	
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
        Dim inputDataSheet, printSheet As Worksheet
        Set inputDataSheet = NatesEvolution.Worksheets("Nate's Evolution")
        Set printSheet = NatesEvolution.Worksheets("Debug")

    ''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
        Dim inputRow, printRow As Integer
        printRow = 2

    '''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
    Dim inputData() As cJobData
    Dim currentRep  As cJobData

    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim inputDataSize As Long

    'inputDataSize = inputDataSheet.UsedRange.Rows.Count
    inputDataSize = 5456
    ReDim inputData(inputDataSize)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''GET AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      For inputRow = 2 To inputDataSize + 2
            With inputDataSheet
                'Get data from sheet and pass it to new Data object'
                Set currentRep = New cJobData
                    currentRep.Customer    = .Cells(inputRow, customerCol).value
                    currentRep.JobID       = .Cells(inputRow, jobCol).value
                    currentRep.kW          = .Cells(inputRow, kWCol).value
                    currentRep.Status      = .Cells(inputRow, statusCol).value
                    currentRep.SubStatus   = .Cells(inputRow, subStatusCol).value
                    currentRep.CreatedDate = .Cells(inputRow, createdDateCol).value
                    currentRep.RepEmail    = .Cells(inputRow, repEmailCol).value
                    currentRep.setIsDocSigned (.Cells(inputRow, isDocSignedCol).value)
                    currentRep.setIsFinalContract (.Cells(inputRow, isFinalContractCol).value)
                    currentRep.setIsInstall
                    currentRep.setIsCancelled
            End With

            Set inputData(inputRow - 2) = currentRep
        Next inputRow
        printRow = 2
        
        For Each printRep In inputData
            
            With printSheet
                    .Cells(printRow, 1).value  = printRep.Customer
                    .Cells(printRow, 2).value  = printRep.JobID
                    .Cells(printRow, 3).value  = printRep.kW
                    .Cells(printRow, 4).value  = printRep.CreatedDate
                    .Cells(printRow, 5).value  = printRep.Status
                    .Cells(printRow, 6).value  = printRep.SubStatus
                    .Cells(printRow, 7).value  = printRep.RepEmail
                    .Cells(printRow, 8).value  = printRep.IsDocSigned
                    .Cells(printRow, 9).value  = printRep.IsFinalContract
                    .Cells(printRow, 10).value = printRep.IsInstall
                    .Cells(printRow, 11).value = printRep.IsCancelled

            End With
            
            printRow = printRow + 1
        Next


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''PRINT VALUES''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Sub




