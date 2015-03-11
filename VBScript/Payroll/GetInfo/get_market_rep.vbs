Function getMarketRep(ByRef currentRep As cRepData, ByVal workBookName As String) As cRepData
   
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    ''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim IDCol, RateCol, StartCol, EndCol As Integer

        IDCol    = 3
        RateCol  = 4
        StartCol = 5
        EndCol   = 6

    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)
        
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim marketDataSheet As Worksheet
        Set marketDataSheet = NatesEvolution.Worksheets("Marketing")

    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim marketDataSize As Long
        marketDataSize = marketDataSheet.Cells(3,IDCol).End(xlDown).Row


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''LOGIC AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    For marketRow = 4 To marketDataSize - 1
    
        With marketDataSheet
            
            If .Cells(marketRow, IDCol).Value = currentRep.ID Then
                currentRep.MarketingRate  = .Cells(marketRow, RateCol).Value
                currentRep.MarkStartDate = .Cells(marketRow, StartCol).Value
                currentRep.MarkEndDate   = .Cells(marketRow, EndCol).Value
                currentRep.setIsMarketing
            End if


        End With
    Next marketRow

    Set getMarketRep = currentRep
End Function