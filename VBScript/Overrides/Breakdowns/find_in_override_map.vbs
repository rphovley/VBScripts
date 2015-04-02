'SORT PAYMENTS BY JOBID, AND THEN OVERRIDE ID'
Function findInOverrideMap(ByVal jobID As String, ByVal uplineRepName As String, ByVal uplineRepType As String) As Currency
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
    
    Dim overrideRepCol, jobCol, overrideTypeCol, amountCol As Integer

    overrideRepCol = 2
    jobCol = 7
    overrideTypeCol = 13
    amountCol = 16

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
        workBookName3 = "February Override Master" & ".xlsx"
    'workBookName = InputBox("What is the master report's name?") & ".xlsm"
    Dim OverrideMaster As Workbook
    Set OverrideMaster = Workbooks(workBookName3)


''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim masterInput As Worksheet
        Set masterInput = MasterReport.Worksheets("Current Data")
        Set masterTest = MasterReport.Worksheets("Test List")

    Dim breakMaster As Worksheet
        Set breakMaster = Breakdown.Worksheets("Master")


    Dim overridePayments As Worksheet
        Set overridePayments = OverrideMaster.Worksheets("Payments")


''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''

Dim thisJobRow, jobRow As Integer


'''''''''''''''''''''''''''''Input Object''''''''''''''''''''''

Dim currentJob As cJobData

'''''''''Found Job in override payments''''''''''''''''''''''''''''''
Dim isJobFound As Boolean
    isJobFound = True

'''''''''Total Payment made for this job''''''''''''''''''''''''''''''
Dim totalPaid As Currency
    totalPaid = 0
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''FIND AND GATHER DATA''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Find job'
On Error GoTo jobIdNotFound:
    thisJobRow = Application.WorksheetFunction.Match(jobID, overidePayments.Range("G:G"), 0)
    jobRow = thisJobRow

'If j'
If isJobFound Then
    'Find upline rep for the job'
    Do Until (jobID <> overridePayments.Cells(jobRow, jobCol).value)
        
        'is this the rep that the override is related to'
        If uplineRepName = overridePayments.Cells(jobRow, overrideRepCol).value Then
            'Is this the matching override Type?'
            If uplineRepType = overridePayments.Cells(jobRow, overrideTypeCol).value Then
                totalPaid = totalPaid + overridePayments.Cells(jobRow, amountCol).value
            End If
        End If

        jobRow = jobRow + 1
    Loop


End If

    findInOverrideMap = totalPaid
jobIdNotFound:
    isJobFound = False
    Resume Next

End Function

