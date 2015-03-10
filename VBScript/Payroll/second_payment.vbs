function secondPayment(ByRef currentJob As cJobData, ByRef currentRep As cRepData, ByVal WorkBookName As String) As cJobData

'Happens after final contract is signed
dim second_payment_rate as currency
	second_payment_rate = 50
dim number_of_kW as integer
dim second_payment_total as currency

'Need to find the total number of kW that qualify for this week's second payment.
second_payment_total = number_of_kW * second_payment_rate

'determine if job is at second payment status and that it hasn't been cancelled'
If currentJob.isFinalContract AND NOT currentJob.isCancelled Then
	currentJob.ThisWeeksSecondPayment = second_payment_total
End If

	Set secondPayment = currentJob

End function