'''''''''''''''''''''''''''''objects'''''''''''''''''''''''''
	Dim job       As cJobData
	Dim rep       As cRepData

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
					Set job = firstPayment(job, rep, workBookName)	
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

'returns the rep object associated with the job'-
Function findRep( ByRef repData As Collection, ByVal repEmail As String) As Integer
		
	For repIndex = 1 To repData.Count
        If repData(repIndex) = repEmail Then
        	findRep = repIndex
        End IF
    Next

End Function

'Returns the scale object associated with the rep'
Function findScale(ByRef scaleData As Collection, ByVal scaleID As Integer) As cScaleData
		
	For Each payScale In scaleData
        If payScale.ID = scaleID Then
        	Set findScale = payScale
        End IF
    Next

End Function

'Returns the slider object associated with the scale ID'
Function findSlider(ByRef sliderData As Collection, ByRef scaleID As Integer) As cSliderData
		
	For Each slider In sliderData
        If slider.ID = scaleID Then
        	Set findSlider = slider
        End IF
    Next

End Function