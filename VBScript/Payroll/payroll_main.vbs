sub payroll_main()

	'This is the main sub that should call everything and have other subs return calculated items
	'It should be the framework or skeleton for the entire program

	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
	'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
	    Dim workBookName As String
	    Dim testName As String
	        testName = "VBA Triforce (Ezra)"
	        workBookName = testName & ".xlsm"
	        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
	    Dim NatesEvolution As Workbook
	        Set NatesEvolution = Workbooks(workBookName)

		'''''''''''''''''''''''''''''Input Array Object''''''''''''''''''''''
	    Dim jobData()  As cJobData
	    Dim repData    As Collection
	    Dim scaleData  As Collection
	    Dim sliderData As Collection

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''MAIN METHODS AND LOGIC'''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	'Load up cJobData array with information from jobs in Nate'sEvolution'
		jobData()      = getjobData(workBookName)
	'Load up repData with Rep Information'
		Set repData    = getRepData(workBookName)
	'Load up scale data with Scale information'
		Set scaleData  = getScaleData(workBookName)
	'Load up slider data with Slider Information'
		Set sliderData = getSliderData(workBookName)
		
	'Get relevant payment information from the payment tabs and update jobData'
		jobData() = getPaymentInfo(jobData, workBookName)
	
	'Process Payment Info'
		jobData() = processPayment(jobData, repData, scaleData, workBookName)
	'print out to the debug sheet all of the relevant job data'

		'printToDebug jobData, workBookName

		printToDebugRep repData, workBookName

		printAllToDebug jobData, repData, workBookName
	'Loop through pending accounts to grab any info for those jobs based on'

End Sub