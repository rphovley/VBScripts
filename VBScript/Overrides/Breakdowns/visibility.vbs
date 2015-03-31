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

    CustomerCol = 1
    jobCol = 2
    KWCol = 3
    StatusCol = 4
    SubStatusCol = 5
    CreatedDateCol = 7
    FinalCol = 8
    RepEmailCol = 17
    StateCol = 24

''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
    Dim workBookName1 As String
        workBookName1 = "3-16-15 Evolve Master Report (Other)" & ".xlsm"
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
    


''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''

Dim inputRow, printRow, jobRow As Integer


''''''''''''''''''''''''''''''Column Counters''''''''''''''''''''''

Dim inputCol, printCol, jobCol As Integer


'''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
Dim currentRep As cRepData

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
    With masterInput
        Set currentJob = New cJobData
            currentJob.Customer    = .Cells(inputRow, CustomerCol).value
            currentJob.JobID       = .Cells(inputRow, jobCol).value
            currentJob.kW          = .Cells(inputRow, KWCol).value
            currentJob.Status      = .Cells(inputRow, StatusCol).value
            currentJob.SubStatus   = .Cells(inputRow, SubStatusCol).value
            currentJob.CreatedDate = .Cells(inputRow, CreatedDateCol).value
            currentJob.repEmail    = .Cells(inputRow, RepEmailCol).value
            currentJob.States      = .Cells(inputRow, StateCol).value
    End With

    jobData.Add currentJob.JobID, currentJob

    inputRow = inputRow + 1

    
    
    'Get rep's name from email
    Dim repRange As Range
    Dim repRow As Integer
    Dim repEmail, repName As String
    Dim totalPaid As Currency
    
    With repMaster
        
        repEmail = currentJob.repEmail
        
        repRow = Application.WorksheetFunction.Match(repEmail, .Range("G:G"), 0)
        
        repName = .Cells(repRow, 8).value
    
    End With
    
    
    With OverrideMaster.Worksheets(MonthName(Month(currentJob.CreatedDate), False) & " " & Year(currentJob.CreatedDate) & " Map")
    
        repRow = Application.WorksheetFunction.Match(repName, .Range("A:A"), 0)
        
        inputCol = 3
        
        'Go through override map to find individual uplines'
        Do Until IsEmpty(.Cells(repRow, inputCol))
            overrideType = .Cells(repRow, inputCol).value
            overrideName = .Cells(repRow, inputCol + 1).value
            ID           = .Cells(repRow, inputCol + 2).value
            overrideRate = .Cells(repRow, inputCol + 3).value           
            
            totalPaid = findInOverrideMap(currentJob.JobID, overrideName, overrideType)
            
            inputCol = inputCol + 4
            

            'Print out to breakdown '
             printToBreakDown currentJob, overrideType, overrideName, totalPaid, repName   
        Loop
    
    
    End With

Loop


''''''''''''''''''TURN ON SCREEN UPDATING''''''''''''''''''''''
Application.ScreenUpdating = True
End Sub

Sub printToBreakDown(ByRef currentJob As cJobData, ByVal overrideType As String, ByVal overrideName As String, ByVal totalPaid As Currency, ByVal repName As String, ByVal overrideRate As Integer)


''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    
    Dim repCol, customerCol, jobIDCol, statusCol, subStatusCol, overrideType as String
    Dim kW_RateCol As Integer
    Dim totalPaidCol as Currency
    Dim kWCol as Double
    Dim dateCreatCol as Date


    repCol       = 1
    customerCol  = 2
    kWCol        = 3
    kW_RateCol   = 4
    totalPaidCol = 5
    dateCreatCol = 6
    jobIDCol     = 7
    statusCol    = 8
    subStatusCol = 9
    overrideType = 10
    

''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''

Dim inputRow, printRow, jobRow As Integer


''''''''''''''''''''''''''''''Column Counters''''''''''''''''''''''

Dim inputCol, printCol, jobCol As Integer


''''''''''''''''''TURN OFF SCREEN UPDATING''''''''''''''''''''''
    Application.ScreenUpdating = False
    
    printRow = Pre-Breakdown.Worksheets(overrideName).Cells(3,2).End(xlDown).End(xlDown).End(xlDown).Row + 1

    With Pre-Breakdown.Worksheets(overrideName)
        .Cells(printRow, repCol).value       = repName
        .Cells(printRow, customerCol).value  = currentJob.Customer
        .Cells(printRow, kWCol).value        = currentJob.kW
        .Cells(printRow, kW_RateCol).value   = overrideRate
        .Cells(printRow, totalPaidCol).value = totalPaid
        .Cells(printRow, dateCreatCol).value = currentJob.CreatedDate
        .Cells(printRow, jobIDCol).value     = currentJob.JobID
        .Cells(printRow, statusCol).value    = currentJob.Status
        .Cells(printRow, subStatusCol).value = currentJob.SubStatus
        .Cells(printRow, overrideType).value = overrideType

    End With
    



''''''''''''''''''TURN ON SCREEN UPDATING''''''''''''''''''''''
    Application.ScreenUpdating = True

End Sub


