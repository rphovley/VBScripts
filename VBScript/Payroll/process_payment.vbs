Function processPayment(ByRef jobData() As cJobData, ByRef repData As Collection, ByRef scaleData As Collection, ByRef sliderData As Collection, ByVal workBookName As String) As cJobData()

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''INIT VARS'''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

	''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
	    Dim NatesEvolution As Workbook
	        Set NatesEvolution = Workbooks(workBookName)

	'''''''''''''''''''''''''''''objects'''''''''''''''''''''''''
	Dim job       As cJobData
	Dim rep       As cRepData
	Dim payScale  As cScaleData
	Dim slider    As cSliderData

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''PROCESS PAYMENT LOOP''''''''''''''''''''''''''''''''''''
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	

	    For jobIndex = 0 To UBound(jobData)
	'''''''''''''''''''''''''''''''''''''THESE OBJECTS HAVE EVERYTHING WE NEED TO PROCESS PAYMENT''''''''''''''
	    	Set job       = jobData(jobIndex)
	    	Set rep       = findRep(repData, job.repEmail)
	    	Set payScale  = findScale(scaleData, rep.PayScaleID)
	    	Set slider    = findScale(sliderData, rep.PayScaleID)
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

					End IF


				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				''''''''''''''''''''''''SECOND PAYMENT''''''''''''''''''''''''''''''
				''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
				Set job = secondPayment(job, rep, workBookName)

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

'returns the rep object associated with the job'
Function findRep( ByRef repData As Collection, ByVal repEmail As String) As cRepData
		
	For Each rep In repData
        If rep.Email = repEmail Then
        	Set findRep = rep
        End IF
    Next

End Function

'Returns the scale object associated with the rep'
Function findScale(ByRef scaleData As CollectionByVal scaleID As Integer) As cScaleData
		
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