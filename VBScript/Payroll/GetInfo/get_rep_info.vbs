Function getRepData(ByVal workBookName As String) As Collection

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Dim repListNameCol, repListEmailCol, repListIDCol, _
        repListScaleCol, repListBlackCol, repListInactiveCol As Integer

        repListIDCol       = 1
        repListEmailCol    = 2
        repListNameCol     = 3
        repListScaleCol    = 4
        repListBlackCol    = 5
        repListInactiveCol = 6

    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

    ''''''''''''''''''''''''''Create Nate's Evolution'''''''''''''
    	createNatesEvo(workBookName)
    	
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim repDataSheet As Worksheet
        Set repDataSheet = NatesEvolution.Worksheets("RepData")

    ''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
    Dim inputRow, printRow, repRow As Integer
        printRow = 2

    '''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
    Dim repData As Collection
        Set repData = New Collection
    Dim currentRep As cRepData

    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim repDataSize As Long
    	repDataSize = repDataSheet.Cells(1,1).End(xlDown).Row - 1

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''GET AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      For inputRow = 2 To repDataSize + 1
            With repDataSheet
                    'Set Values for object from the Rep list'
	                Set currentRep = New cRepData             
                        currentRep.ID       = .Cells(inputRow, repListIDCol).Value
                        currentRep.Email    = .Cells(inputRow, repListEmailCol).Value
                        currentRep.Name     = .Cells(inputRow, repListNameCol).Value
                        currentRep.PayScaleID  = .Cells(inputRow, repListScaleCol).Value
                        currentRep.setIsBlackList(.Cells(inputRow, repListBlackCol).Value)
                        currentRep.setIsInactive(Cells(inputRow, repListInactiveCol).Value)
	                    
	                ''''''''''Add currentRep to the jobData Collection''''''''''''
	                repData.Add currentRep, currentRep.Email
            End With

            
        Next inputRow

        Set getRepData = repData
        
        


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''PRINT VALUES''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
End Function




