Function getSliderData(ByVal workBookName As String) As Collection

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Dim sliderIDCol, quartileMaxCol, quartileKWCol, _
        quartileOneRow, quartileTwoRow, quartileThreeRow, quratileFourRow As Integer

        sliderIDCol      = 5
        quartileMaxCol   = 7
        quartileKWCol    = 8
        quartileOneRow   = 1
        quartileTwoRow   = 2
        quartileThreeRow = 3
        quartileFourRow  = 4


    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master sliderort's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)
    	
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim sliderDataSheet As Worksheet
        Set sliderDataSheet = NatesEvolution.Worksheets("Payscales")


    '''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
    Dim sliderData As Collection
        Set sliderData = New Collection
    Dim slider As cSliderData
    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim sliderDataSize As Long
    	sliderDataSize = sliderDataSheet.Cells(1,1).End(xlDown).Row - 1


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''GET AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      For inputRow = 1 To (sliderDataSize / 4) + 1
            With sliderDataSheet
                    'Set Values for object from the slider list'
	                Set slider = New cSliderData 

                        slider.ID    = .Cells(1 + quartileOneRow, sliderIDCol).Value

                        'quartile MAX'
                        slider.FirstQuartileMax  = .Cells(inputRow + quartileOneRow, quartileMaxCol).Value
                        slider.SecondQuartileMax = .Cells(inputRow + quartileTwoRow, quartileMaxCol).Value
                        slider.ThirdQuartileMax  = .Cells(inputRow + quartileThreeRow, quartileMaxCol).Value
                        slider.FourthQuartileMax = .Cells(inputRow + quartileFourRow, quartileMaxCol).Value

                        'quartile kW'   
                        slider.FirstQuartileKW  = .Cells(inputRow + quartileOneRow, quartileKWCol).Value
                        slider.SecondQuartileKW = .Cells(inputRow + quartileTwoRow, quartileKWCol).Value
                        slider.ThirdQuartileKW  = .Cells(inputRow + quartileThreeRow, quartileKWCol).Value
                        slider.FourthQuartileKW = .Cells(inputRow + quartileFourRow, quartileMaxCol).Value  

                         
            End With


            ''''''''''Add currentslider to the jobData Collection''''''''''''
            sliderData.Add slider, str(slider.ID)

            inputRow = inputRow + 3
        Next inputRow

       Set getPaysliderData = sliderData
        
End Function