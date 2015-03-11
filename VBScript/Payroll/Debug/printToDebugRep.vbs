Sub printToDebugRep(ByRef repData As Collection, ByVal workBookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim repListNameCol, repListEmailCol, repListIDCol, _
        repListScaleCol, repListBlackCol, repListInactiveCol, _
        repListIsNewRep, repListIsSliderCol, repSliderDateCol As Integer

        repListIDCol       = 1
        repListEmailCol    = 2
        repListNameCol     = 3
        repListScaleCol    = 4
        repListBlackCol    = 5
        repListInactiveCol = 6
        repListIsNewRep    = 7
        repListIsSliderCol = 8
        repSliderDateCol   = 9
        
''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim printSheet As Worksheet
        Set printSheet = NatesEvolution.Worksheets("Debug Rep")

''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
    Dim printRow As Integer
        printRow = 2

    Const EMPTYDATE = #12:00:00 AM#
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''PRINT REPS TO DEBUG SHEET'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For index = 2 To repData.Count + 1
            
            With printSheet
            
                    .Cells(index, repListIDCol).value          = repData(index - 1).ID
                    .Cells(index, repListEmailCol).value       = repData(index - 1).Email
                    .Cells(index, repListNameCol).value        = repData(index - 1).Name
                    .Cells(index, repListScaleCol).value       = repData(index - 1).PayScaleID
                    .Cells(index, repListBlackCol).value       = repData(index - 1).IsBlackList
                    .Cells(index, repListInactiveCol).value    = repData(index - 1).IsInactive
                    .Cells(index, repListIsSliderCol).value    = repData(index - 1).IsSlider
                    if repData(index - 1).IsSlider = True Then
                        .Cells(index, repSliderDateCol).value    = repData(index - 1).StartSliderDate
                    End if
            End With
            
        printRow = printRow + 1
    Next index

End Sub


