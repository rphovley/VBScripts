Sub getWeatherData(ByVal workBookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        
    Dim repCol, idCol, emailCol As Integer
        repCol   = 1
        idCol   = 2
        emailCol = 3


    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master weatherort's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)
    	
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim weatherDataSheet As Worksheet
        Set weatherDataSheet = NatesEvolution.Worksheets("Weather")


    '''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
    Dim weather As cWeatherData
    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim weatherDataSize As Long
    	weatherDataSize = weatherDataSheet.Cells(1,idCol).End(xlDown).Row

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''GET AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

      For inputRow = 2 To weatherDataSize
            With weatherDataSheet
                    'Set Values for object from the weather list'
	                Set weather = New cWeatherData 
                        weather.Name  = .Cells(inputRow, repCol).Value
                        weather.ID    = .Cells(inputRow, idCol).Value
                        weather.Email = .Cells(inputRow, emailCol).Value
                         
            End With


            ''''''''''Add currentweather to the jobData Collection''''''''''''
            payroll_main.weatherData.Add  weather.Email, weather

        Next inputRow
        
End Sub