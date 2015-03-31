testFindOverride()
''''''''''''''''''''''''''Input Array Object''''''''''''''''''''''
Dim repData     As Scripting.Dictionary
Dim jobData     As Scripting.Dictionary

''''''''''''''''''''''''''''Set objects'''''''''''''''''''''''
Set repData     = New Scripting.Dictionary
Set jobData     = New Scripting.Dictionary

''''''''''''''''''TURN OFF SCREEN UPDATING''''''''''''''''''''''
application.screenupdating = False
    
''''''''''''''''''''''''''''''Columns''''''''''''''''''''''

Dim masCustomerCol, masJobCol, masKWCol, masStatusCol, masSubStatusCol, _
masDateCol, masFinalCol, masRepEmailCol, masStateCol As Integer

    masCustomerCol        = 1   
    masJobCol             = 2
    masKWCol              = 3
    masStatusCol          = 4
    masSubStatusCol       = 5
    masCreatedDateCol     = 7
    masFinalCol           = 8
    masRepEmailCol        = 17
    masStateCol           = 24


''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
Dim workBookName1 As String
    workBookName1 = "3-16-15 Evolve Master Report" & ".xlsx"
'workBookName = InputBox("What is the master report's name?") & ".xlsm"       
Dim MasterReport As Workbook
    Set MasterReport = Workbooks(workBookName1)


''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
Dim masterInput, masterTest As Worksheet
    masterInput = MasterReport.Worksheets("Current Data")
    masterTest = MasterReport.Worksheets("Test List")


''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''

Dim jobRow As Integer
    jobRow = 2

'''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
Dim currentJob As cJobData

'''''''''Found Job in override payments''''''''''''''''''''''''''''''
Dim isJobFound As Boolean
    isJobFound = True

'''''''''Total Payment made for this job''''''''''''''''''''''''''''''
Dim totalPaid As Currency
    totalPaid = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''TEST OUT FOR ALL JOBS'''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

With masterInput
    Do Until(IsEmpty(.Cells(jobRow, 1)))

        masterTest.Cells(jobRow, 2).value = findInOverrideMap(currentJob, "Matt Ingalls", "RC")
        masterTest.Cells(jobRow, 3).value = findInOverrideMap(currentJob, "Matt Ingalls", "M")
        masterTest.Cells(jobRow, 4).value = findInOverrideMap(currentJob, "Matt Ingalls", "RG")
        masterTest.Cells(jobRow, 5).value = findInOverrideMap(currentJob, "Matt Ingalls", "D")

        jobRow = jobRow + 1
    Loop
End With
End Sub