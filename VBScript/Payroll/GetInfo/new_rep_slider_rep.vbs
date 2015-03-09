Function newRepSliderRep(ByRef currentRep As cRepData, ByVal workBookName As String) As cRepData
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim masCustomerCol, masJobCol, masKWCol, masStatusCol, masSubStatusCol, _
        masDateCol, masRepEmailCol As Integer

        masCustomerCol        = 1   
        masJobCol             = 2
        masKWCol              = 3
        masStatusCol          = 4
        masSubStatusCol       = 5
        masCreatedDateCol     = 7
        masRepEmailCol        = 17

    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)
        
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim jobDataSheet As Worksheet
        Set jobDataSheet = NatesEvolution.Worksheets("Master Input")

    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim jobDataSize As Long
        jobDataSize = jobDataSheet.Cells(1,1).End(xlDown).Row - 1


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''LOGIC AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For jobRow = 2 To jobDataSize - 1
    
        With jobDataSheet
            
            If .Cells(jobRow, 17).Value = currentRep.Email Then

                'Determine if the rep is new or not'
                If DateDiff("d", .Cells(jobRow, 7).value, Now()) >= 60 And currentRep.IsNewRep Then                        
                    currentRep.IsNewRep = False
                End If

                'Code for determining if rep is on sliding pay scale or not'
                currentRep.KwSum = currentRep.KwSum + .Cells(jobRow,3).Value
                
                If currentRep.KwSum > 300 And currentRep.IsSlider = False Then
                    currentRep.IsSlider = True
                    currentRep.StartSliderDate = .Cells(jobRow,7).Value
                End If
            End if


        End With
    Next jobRow

    Set newRepSliderRep = currentRep
End Function