Sub installs(ByRef currentJob As cJobData, ByRef currentRep as cRepData, ByVal WorkBookName As String)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim customerCol, jobCol, kWCol, createdDateCol, _
        paymentAmountCol, paymentDateCol, _
        secondPaymentCol, secondDateCol, _
        finalPaymentCol, finalDateCol As Integer

        customerCol      = 3
        jobCol           = 4
        kWCol            = 5
        createdDateCol   = 6
        paymentAmountCol = 7
        paymentDateCol   = 8
        secondPaymentCol = 9
        secondDateCol    = 10
        finalPaymentCol  = 12
        finalDateCol     = 13
        
''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim InstalledSheet As Worksheet
        Set InstalledSheet = NatesEvolution.Worksheets("Installed")

''''''''''''''''''''''Set PayScales Objects''''''''''''''''''
	Dim scales as cScaleData
	Dim slider As cSliderData
''''''''''''''''''''''''Rate used to calculate pay'''''''''''
	Dim Rate as Currency

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''GET DOWN TO BUSINESS''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'If the job is installed, it should be paid in full'
If currentJob.IsInstall Then
	'Determine what it means for this rep to be "Paid in full"'
	'if the rep is not marketing'
	If NOT currentRep.IsMarketing Then
		Set scales = payroll_main.scaleData(CStr(currentRep.PayScaleID))
		If NOT scales IS Nothing AND scales.Rate = "Sliding" Then
			Set slider = payroll_main.sliderData(currentRep.PayScaleID)
		End If

		'If they aren't on a sliding pay scale'
		If slider IS NOTHING THen
			'Full payment equals the remainder of payment on "kW * the_reps_rate"'
			Rate = scales.Rate			
			'Increment our install pool for this rep'
			'Print out to Installed and remove any instances of job from 1st payments and second payments and cancellations'
		'If they are on a sliding pay scale'
		Else
			'NOTE: Sliding pay is paid out based on how many kW they sold during the month the job was sold.  They are paid up front at #$200 and then the amount is'
			'adjusted during the next month. This only handles the weekly calculation based on the flat $200.  Another script will be run monthly to calculate whatever'
			'increase in pay they get based on the sliding pay'
			'If the rep is on a sliding pay scale, we first will determine if the current job was '
			Rate = 200
		End If

	'If they are a marketing rep'
	Else
		Rate = currentRep.MarketingRate
	End If

	currentJob.ThisWeekFinalPayment = (currentJob.kW * Rate) - currentJob.WhatWasPaid
	currentJob.FinalPaymentAmount = currentJob.ThisWeekFinalPayment
	'Increment our install pool for this rep'
	currentRep.InstallPool = currentRep.InstallPool + currentJob.FinalPaymentAmount
	currentJob.FinalPaymentDate = Date()
	currentJob.setWhatWasPaid
	
	
	'Print out to Installed and remove any instances of job from 1st payments and second payments and cancellations'
	printInstallation currentJob, currentRep, Rate, WorkBookName

	'Remove from first payments'
	'Remove from second payments'
	'Remove from cancellations'

End If

End Sub

Sub printInstallation(ByRef currentJob As cJobData, ByRef currentRep As cRepData, ByVal Rate As Currency, ByVal WorkBookName As String)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


''''''''''''''''''''''''''''''Columns''''''''''''''''''''''
    Dim repCol, repIDCol, customerCol, jobCol, kWCol, createdDateCol, _
        paymentAmountCol, paymentDateCol, _
        secondPaymentCol, secondDateCol, rateCol, _
        finalPaymentCol, finalDateCol As Integer

        repCol           = 1
        repIDCol         = 2
        customerCol      = 3
        jobCol           = 4
        kWCol            = 5
        createdDateCol   = 6
        paymentAmountCol = 7
        paymentDateCol   = 8
        secondPaymentCol = 9
        secondDateCol    = 10
        rateCol	         = 11
        finalPaymentCol  = 12
        finalDateCol     = 13
        
''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)
	
	'worksheets'
	Dim Installed As Worksheet
		Set Installed = NatesEvolution.Worksheets("Installed")
	'Row'
	Dim printRow As Integer
		printRow = Installed.Cells(1,1).End(xlDown).Row + 1

	With Installed
		.Cells(printRow, repCol)           = currentRep.Name
		.Cells(printRow, repIDCol)         = currentRep.ID
		.Cells(printRow, customerCol)      = currentJob.Customer
		.Cells(printRow, JobID)            = currentJob.JobID
		.Cells(printRow, kWCol)            = currentJob.kW
		.Cells(printRow, createdDateCol)   = currentJob.CreatedDate
		.Cells(printRow, paymentAmountCol) = currentJob.FirstPaymentAmount
		.Cells(printRow, paymentDateCol)   = currentJob.FirstPaymentDate
		.Cells(printRow, secondPaymentCol) = currentJob.SecondPaymentAmount
		.Cells(printRow, secondDateCol)    = currentJob.SecondPaymentDate
		.Cells(printRow, rateCol)          = Rate
		.Cells(printRow, finalPaymentCol)  = currentJob.ThisWeekFinalPayment
		.Cells(printRow, finalDateCol)     = Date()
	End With

End Sub