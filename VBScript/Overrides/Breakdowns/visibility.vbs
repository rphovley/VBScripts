Sub visibility()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''



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



    Dim workBookName2 As String
        workBookName2 = "Pre-Breakdown" & ".xlsx"
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
        masterInput = MasterReport.Worksheets("Current Data")

    Dim breakMaster As Worksheet
        breakMaster = Breakdown.Worksheets("Master")


''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''

'Dim inputRow, printRow, jobRow, repRow As Integer'


'''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
Dim currentRep As cRepData

Dim currentJob As cJobData


'''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''''''''''''''


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''LOGIC ''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'This gets all of the data needed from the Evolve Master Report workbook, Current data worksheet'
do until isEmpty(ActiveCell.Value)
    with masterInput
        .Cells(1,2).Activate

        Set currentJob = New cJobData
            currentJob.Customer     = ActiveCell.Value 'masCustomerCol'
            currentJob.JobID        = ActiveCell.offset(0, masJobCol-1).Value
            currentJob.kW           = ActiveCell.offset(0, masKWCol-1).Value
            currentJob.Status       = ActiveCell.offset(0, masStatusCol-1).Value
            currentJob.SubStatus    = ActiveCell.offset(0, masSubStatusCol-1).Value
            currentJob.CreatedDate  = ActiveCell.offset(0, masCreatedDateCol-1).Value
            currentJob.RepEmail     = ActiveCell.offset(0, masRepEmailCol-1).Value
            currentJob.States       = ActiveCell.offset(0, masStateCol-1).Value

    end with

    jobData.Add currentJob.JobID, currentJob

    ActiveCell.offset(1,0).Activate

loop
















End Sub
