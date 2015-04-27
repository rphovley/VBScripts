Function newRepSliderRep(ByRef currentRep As cRepData, ByVal workBookName As String) As cRepData
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim masCustomerCol, masJobCol, masKWCol, masStatusCol, masCreatedDateCol, masSubStatusCol, _
        masDateCol, masRepEmailCol, mainmenuDatesCol As Integer

        masCustomerCol        = 1   
        masJobCol             = 2
        masKWCol              = 3
        masStatusCol          = 4
        masSubStatusCol       = 5
        masCreatedDateCol     = 7
        masRepEmailCol        = 17
        mainmenuDatesCol      = 3
		
	''''''''''''''''''''''''Rows'''''''''''''''''''''''''''
		mainmenuStartDateRow = 5
		mainmenuEndDateRow = 6

    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)
        
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim jobDataSheet, MainMenuSheet As Worksheet
        Set jobDataSheet = NatesEvolution.Worksheets("Master Input")
		Set MainMenuSheet = NatesEvolution.Worksheets("Main Menu")

    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim jobDataSize As Long
        jobDataSize = jobDataSheet.Cells(1,1).End(xlDown).Row - 1

    ''''''''''''''''''''''''''''First Sale Boolean'''''''''''''''''''
    Dim isFirstSale As Boolean
        isFirstSale = True
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''LOGIC AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''''''''''''''''''''''''''Sets start and end dates for pay period'''''''''''''''''''''''''''''''''''
Dim PayPeriodStartDate, PayPeriodEndDate as Date
	PayPeriodStartDate = MainMenuSheet.cells(mainmenuStartDateRow, mainmenuDatesCol).value
	PayPeriodEndDate = MainMenuSheet.cells(mainmenuEndDateRow, mainmenuDatesCol).value
	
    For jobRow = 2 To jobDataSize - 1
    
        With jobDataSheet
            
            If .Cells(jobRow, masRepEmailCol).Value = currentRep.Email Then

                'Code for determining if rep is on sliding pay scale or not'
                currentRep.KwSum = currentRep.KwSum + .Cells(jobRow,3).Value
                
                If currentRep.KwSum > 300 And currentRep.IsSlider = False Then
                    currentRep.IsSlider = True
                    currentRep.StartSliderDate = .Cells(jobRow,7).Value
                End If

                'Code for getting first job date for the rep'
                If isFirstSale Then
                    currentRep.FirstJobDate = .Cells(jobRow, masCreatedDateCol).Value
                    isFirstSale = False
                End If
				
            End if

        End With
    Next jobRow

    Set newRepSliderRep = currentRep
End Function