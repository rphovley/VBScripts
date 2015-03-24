'''''''''''''''''''''''''''''objects'''''''''''''''''''''''''
	Dim job       As cJobData
	Dim rep       As cRepData
	Dim weather   As cWeatherData

Sub processPayment(ByVal workBookName As String)

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''INIT VARS'''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
	    Dim NatesEvolution As Workbook
	        Set NatesEvolution = Workbooks(workBookName)

	'''''''''''''''''''''''''''''REP INDEX''''''''''''''''''''''''''''''
		Dim repIndex As Integer

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''PROCESS PAYMENT LOOP''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	

	    For jobIndex = 0 To UBound(payroll_main.jobData)
	'''''''''''''''''''''''''''''''''''''THESE OBJECTS HAVE EVERYTHING WE NEED TO PROCESS PAYMENT''''''''''''''
	    	Set job = payroll_main.jobData(jobIndex)
	    	On Error Resume Next
	    	Set rep = payroll_main.repData(job.repEmail)
	    	Set weather = payroll_main.weatherData(job.repEmail)

	'''''''''''''''''''''''''''''''''''''THESE OBJECTS HAVE EVERYTHING WE NEED TO PROCESS PAYMENT''''''''''''''


	    	'IGNORE ANY JOBS THAT HAVE ALREADY BEEN PAID OUT IN FULL AND HAVE A VALID REP'
	    	If job.FinalPaymentAmount = 0 AND Not rep is nothing Then

	    		'IGNORE ANY JOBS THAT HAVE ALREADY BEEN PAID SECOND PAYMENTS'
	    		If job.SecondPaymentAmount = 0 Then

	    			'IGNORE ANY JOBS THAT HAVE ALREADY BEEN PAID FIRST PAYMENTS'
	    			If job.FirstPaymentAmount = 0 Then
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					''''''''''''''''''''''''FIRST PAYMENT'''''''''''''''''''''''''''''''
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					Set job = firstPayment(rep, job, weather, workBookName)	
					End IF


				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				''''''''''''''''''''''''SECOND PAYMENT''''''''''''''''''''''''''''''
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				Set job = secondPayment(job, rep, workBookName)

				End If

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''FINAL PAYMENT'''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			installs job, rep, workBookName



			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''CANCELLATIONS'''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			Set job = cancells(job, rep, workBookName)

			End If
			payroll_main.repData.Remove rep.Email
			payroll_main.repData.Add rep.Email, rep
			Set payroll_main.jobData(jobIndex) = job
		Next jobIndex
End Sub