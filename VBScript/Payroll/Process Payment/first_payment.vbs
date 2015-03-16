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
 Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
        createdDateCol, repEmailCol, isDocSignedCol, isFinalContractCol As Integer

        customerCol = 1
        jobCol = 2
        createdDateCol = 6
        repEmailCol = 9
        
 ''''''''''''''''''''''''''''''Calculates the first payment''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FOR ACCOUNTS LESS THAN 60 DAYS''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
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
    
    Set firstPayment = currentJob
    
End Function
        
        
        
Sub printFirst(ByRef currentRep As cRepData)


    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(WorkBookName)

    ''''''''''''''''''''''''''Create Nate's Evolution'''''''''''''
        createNatesEvo (WorkBookName)
        
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim jobDataSheet, printSheet As Worksheet
        Set jobDataSheet = NatesEvolution.Worksheets("Nate's Evolution")
        Set printSheet = NatesEvolution.Worksheets("Debug")

    ''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
    Dim inputRow, printRow As Integer
        printRow = 2

    '''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
    Dim jobData() As cJobData
    Dim currentJob  As cJobData

    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim jobDataSize As Long
        jobDataSize = jobDataSheet.Cells(1, 1).End(xlDown).row - 2
        ReDim jobData(jobDataSize)

End Sub

