Sub getPayScaleData(ByVal workBookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Dim scaleIDCol, scaleNameCol, scaleRateCol As Integer

        scaleIDCol   = 1
        scaleNameCol = 2
        scaleRateCol = 3


    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master scaleort's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)
    	
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim scaleDataSheet As Worksheet
        Set scaleDataSheet = NatesEvolution.Worksheets("Payscales")


    '''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
    Dim payScale As cScaleData

    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim scaleDataSize As Long
    	scaleDataSize = scaleDataSheet.Cells(1,1).End(xlDown).Row - 1


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''GET AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      For inputRow = 2 To scaleDataSize + 1
            With scaleDataSheet
                    'Set Values for object from the scale list'
	                Set payScale = New cScaleData 

                        payScale.ID    = .Cells(inputRow, scaleIDCol).Value
                        payScale.Name  = .Cells(inputRow, scaleNameCol).Value
                        payScale.Rate  = .Cells(inputRow, scaleRateCol).Value       
            End With


            ''''''''''Add currentscale to the jobData Collection''''''''''''
            payroll_main.scaleData.Add Trim(CStr(payScale.ID)), payScale

        Next inputRow
        
End Sub