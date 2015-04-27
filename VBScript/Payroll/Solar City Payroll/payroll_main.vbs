'''''''''''''''''''''''''''''Input Array Object''''''''''''''''''''''
	    Public jobData()   As cJobData
	    Public repData     As Scripting.Dictionary
	    Public scaleData   As Scripting.Dictionary
	    Public sliderData  As Scripting.Dictionary
	    Public weatherData As Scripting.Dictionary

sub payroll_main()

	'This is the main sub that should call everything and have other subs return calculated items
	'It should be the framework or skeleton for the entire program

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	''''''''''''''''''TURN OFF SCREEN UPDATING''''''''''''''''''''''
	application.screenupdating = False

	''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
	    Dim workBookName As String
	    Dim testName As String
	        testName = "VBA Triforce (Ezra)"
	        workBookName = testName & ".xlsm"
	        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
	    Dim NatesEvolution As Workbook
	        Set NatesEvolution = Workbooks(workBookName)

	''''''''''''''''''''''''''''Set objects'''''''''''''''''''''''
	Set repData     = New Scripting.Dictionary
	Set scaleData   = New Scripting.Dictionary
	Set sliderData  = New Scripting.Dictionary
	Set weatherData = New Scripting.Dictionary

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''MAIN METHODS AND LOGIC'''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	'Get weather Data exceptions'
		getWeatherData workBookName
	'Load up cJobData array with information from jobs in Nate'sEvolution'
		jobData() = getjobData(workBookName)
	'Load up repData with Rep Information'
		getRepData workBookName
	'Load up scale data with Scale information'
		getPayScaleData workBookName
	'Load up slider data with Slider Information'
		getSliderData workBookName		

	'Get relevant payment information from the payment tabs and update jobData'
		getPaymentInfo workBookName

	'Get count information for reps and the jobs they did this past week'
		getCountInfo workBookName

	'Process Payment Info'
		processPayment workBookName 
	'print out to the debug sheet all of the relevant job data'


		'printToDebug jobData, workBookName

		printToDebugRep repData, workBookName

		printAllToDebug workBookName
	'Loop through pending accounts to grab any info for those jobs based on'

	''''''''''''''''''TURN ON SCREEN UPDATING''''''''''''''''''''''
		application.screenupdating = True

End Sub