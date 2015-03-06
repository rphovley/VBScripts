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
	    Dim jobData() As cJobData

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''MAIN METHODS AND LOGIC'''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	'Load up cJobData array with information from jobs in Nate'sEvolution'
		jobData() = getjobData(workBookName)

	'Get relevant payment information from the payment tabs and update jobData'
		jobData() = getPaymentInfo jobData, workBookName

	'print out to the debug sheet all of the relevant job data'
		printToDebug jobData, workBookName

	'Loop through pending accounts to grab any info for those jobs based on'

End Sub