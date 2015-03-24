Sub createNatesEvo(ByVal workBookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
        createdDateCol, repEmailCol, finalContratCol, stateCol As Integer

        customerCol        = 1   
        jobCol             = 2
        kWCol              = 3
        statusCol          = 4
        subStatusCol       = 5
        createdDateCol     = 6
        repEmailCol        = 7
        finalContratCol    = 8
        stateCol           = 9
        
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
        'workBookName = InputBox("What is the master report's name?") & ".xlsm"       
        Dim MasterReport As Workbook
        Set MasterReport = Workbooks(workBookName)

    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
        Dim masterInput As Worksheet
            masterInput = MasterReport.Worksheets("Current Data")

    ''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''


    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''

    
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''LOGIC ''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


End Sub
