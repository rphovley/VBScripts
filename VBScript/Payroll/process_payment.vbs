Function processPayment(ByRef jobData() As cJobData, ByRef repData As Collection, ByVal workBookName As String) As cJobData()

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''INIT VARS'''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
	    Dim NatesEvolution As Workbook
	        Set NatesEvolution = Workbooks(workBookName)

	'''''''''''''''''''''''''''''objects'''''''''''''''''''''''''


	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''PROCESS PAYMENT LOOP''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	Dim job As cJobData

	    For jobIndex = 0 To UBound(jobData)
	    	Set job = jobData(jobIndex)

	    	'IGNORE ANY JOBS THAT HAVE ALREADY BEEN PAID OUT IN FULL'
	    	If job.FinalPaymentAmount = 0 Then

	    		'IGNORE ANY JOBS THAT HAVE ALREADY BEEN PAID SECOND PAYMENTS'
	    		If job.SecondPaymentAmount = 0 Then

	    			'IGNORE ANY JOBS THAT HAVE ALREADY BEEN PAID FIRST PAYMENTS'
	    			If job.FirstPaymentAmount = 0 Then
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
					''''''''''''''''''''''''FIRST PAYMENT'''''''''''''''''''''''''''''''
					''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


					End IF


				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				''''''''''''''''''''''''SECOND PAYMENT''''''''''''''''''''''''''''''
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				Set job = secondPayment(job, workBookName)

				End If

			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''FINAL PAYMENT'''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''




			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			''''''''''''''''''''''''CANCELLATIONS'''''''''''''''''''''''''''''''
			''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
			End If

			Set jobData(jobIndex) = job
		Next jobIndex
	processPayment = jobData
End Function