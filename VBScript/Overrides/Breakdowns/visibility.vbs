Sub visibility()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



''''''''''''''''''''''''''Input Array Object''''''''''''''''''''''
Dim repData     As Scripting.Dictionary
Dim jobData     As Scripting.Dictionary

''''''''''''''''''''''''''''Set objects'''''''''''''''''''''''
Set repData = New Scripting.Dictionary
Set jobData = New Scripting.Dictionary


''''''''''''''''''TURN OFF SCREEN UPDATING''''''''''''''''''''''
Application.ScreenUpdating = False


''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    
    Dim masCustomerCol, masJobCol, masKWCol, masStatusCol, masSubStatusCol, _
    masDateCol, masFinalCol, masRepEmailCol, masStateCol As Integer

    customerCol = 1
    jobCol = 2
    kWCol = 3
    statusCol = 4
    subStatusCol = 5
    CreatedDateCol = 7
    FinalCol = 8
    RepEmailCol = 17
    StateCol = 24

''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
    Dim workBookName1 As String
        workBookName1 = "3-16-15 Evolve Master Report" & ".xlsx"
    'workBookName = InputBox("What is the master report's name?") & ".xlsm"
    Dim MasterReport As Workbook
    Set MasterReport = Workbooks(workBookName1)



    Dim workBookName2 As String
        workBookName2 = "Pre-Breakdown" & ".xlsm"
    'workBookName = InputBox("What is the master report's name?") & ".xlsm"
    Dim Breakdown As Workbook
    Set Breakdown = Workbooks(workBookName2)



    Dim workBookName3 As String
        workBookName3 = "February Override Master" & ".xlsm"
    'workBookName = InputBox("What is the master report's name?") & ".xlsm"
    Dim OverrideMaster As Workbook
    Set OverrideMaster = Workbooks(workBookName3)


''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim masterInput As Worksheet
    Set masterInput = MasterReport.Worksheets("Current Data")

    Dim breakMaster As Worksheet
    Set breakMaster = Breakdown.Worksheets("Master")
    
    Dim repMaster As Worksheet
    Set repMaster = Breakdown.Worksheets("Reps")
    
    Dim repBreakdown As Worksheet

''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''

Dim inputRow, printRow, jobRow As Integer

Dim isRepFound As Boolean
    isRepFound = True
''''''''''''''''''''''''''''''Column Counters''''''''''''''''''''''

Dim inputCol, printCol As Integer


'''''''''''''''''''''''''''''Input Object''''''''''''''''''''''

Dim currentJob As cJobData

'''''''''''''''''''''''''''''Override Data''''''''''''''''''''''

Dim overrideName, overrideType As String
Dim overrideRate As Currency
Dim ID As Integer


'''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''LOGIC ''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

inputRow = 2


'This gets all of the data needed from the Evolve Master Report workbook, Current data worksheet'
Do Until IsEmpty(masterInput.Cells(inputRow, 1))
    DoEvents
    isRepFound = True
    With masterInput
        Set currentJob = New cJobData
            currentJob.Customer = .Cells(inputRow, customerCol).value
            currentJob.jobID = .Cells(inputRow, jobCol).value
            currentJob.kW = .Cells(inputRow, kWCol).value
            currentJob.Status = .Cells(inputRow, statusCol).value
            currentJob.SubStatus = .Cells(inputRow, subStatusCol).value
            currentJob.CreatedDate = .Cells(inputRow, CreatedDateCol).value
            currentJob.RepEmail = .Cells(inputRow, RepEmailCol).value
            currentJob.States = .Cells(inputRow, StateCol).value
    End With

    jobData.Add currentJob.jobID, currentJob

    inputRow = inputRow + 1

    
    
    'Get rep's name from email
    Dim repRange As Range
    Dim repRow As Integer
    Dim RepEmail, repName As String
    Dim totalPaid As Currency
    
    Const MAYDATE = #5/1/2014#
    Const MARCHDATE = #3/1/2015#
    
    With repMaster
        
        RepEmail = currentJob.RepEmail
        
        repRow = Application.WorksheetFunction.Match(RepEmail, .Range("G:G"), 0)
        
        repName = .Cells(repRow, 8).value
    
    End With
    
    
    If currentJob.CreatedDate < MAYDATE Then
        Set repBreakdown = OverrideMaster.Sheets("May 2014 Map")
    ElseIf currentJob.CreatedDate > MARCHDATE Then
        Set repBreakdown = OverrideMaster.Sheets("February 2015 Map")
    Else
        Set repBreakdown = OverrideMaster.Worksheets(MonthName(Month(currentJob.CreatedDate), False) & " " & Year(currentJob.CreatedDate) & " Map")
    End If
    
    With repBreakdown
        On Error GoTo repNotFound:
        repRow = Application.WorksheetFunction.Match(repName, .Range("A:A"), 0)
        
        inputCol = 3
       If isRepFound Then
            'Go through override map to find individual uplines'
            Do Until IsEmpty(.Cells(repRow, inputCol))
                DoEvents
                overrideType = .Cells(repRow, inputCol).value
                overrideName = .Cells(repRow, inputCol + 1).value
                ID = .Cells(repRow, inputCol + 2).value
                overrideRate = .Cells(repRow, inputCol + 3).value
                
                totalPaid = findInOverrideMap(currentJob.jobID, overrideName, overrideType)
                
                inputCol = inputCol + 4
                
    
                'Print out to breakdown '
                 printToBreakDown currentJob, overrideType, overrideName, totalPaid, repName, overrideRate
            Loop
        End If
    
    
    End With

Loop


repNotFound:
    isRepFound = False
Resume Next
''''''''''''''''''TURN ON SCREEN UPDATING''''''''''''''''''''''
Application.ScreenUpdating = True
End Sub

Sub printToBreakDown(ByRef currentJob As cJobData, ByVal overrideType As String, ByVal overrideName As String, ByVal totalPaid As Currency, ByVal repName As String, ByVal overrideRate As Integer)


''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    
    Dim repCol, customerCol, jobIDCol, statusCol, subStatusCol, overrideTypeCol, _
     kW_RateCol, totalPaidCol, kWCol, dateCreatCol As Integer

    repCol = 2
    customerCol = 3
    kWCol = 4
    kW_RateCol = 5
    totalPaidCol = 6
    dateCreatCol = 7
    jobIDCol = 8
    statusCol = 9
    subStatusCol = 10
    overrideTypeCol = 11
    

''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''

Dim inputRow, printRow, jobRow As Integer


''''''''''''''''''''''''''''''Column Counters''''''''''''''''''''''

Dim inputCol, printCol, jobCol As Integer


''''''''''''''''''TURN OFF SCREEN UPDATING''''''''''''''''''''''
    Application.ScreenUpdating = False
    
    printRow = Workbooks("Pre-Breakdown.xlsm").Worksheets(overrideName).Cells(3, 2).End(xlDown).End(xlDown).End(xlDown).Row + 1

    With Workbooks("Pre-Breakdown.xlsm").Worksheets(overrideName)
        .Cells(printRow, repCol).value = repName
        .Cells(printRow, customerCol).value = currentJob.Customer
        .Cells(printRow, kWCol).value = currentJob.kW
        .Cells(printRow, kW_RateCol).value = overrideRate
        .Cells(printRow, totalPaidCol).value = totalPaid
        .Cells(printRow, dateCreatCol).value = currentJob.CreatedDate
        .Cells(printRow, jobIDCol).value = currentJob.jobID
        .Cells(printRow, statusCol).value = currentJob.Status
        .Cells(printRow, subStatusCol).value = currentJob.SubStatus
        .Cells(printRow, overrideTypeCol).value = overrideType

    End With
    



''''''''''''''''''TURN ON SCREEN UPDATING''''''''''''''''''''''
    Application.ScreenUpdating = True

End Sub




