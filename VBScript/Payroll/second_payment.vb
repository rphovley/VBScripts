function second_payment()

'Happens after final contract is signed
dim second_payment_rate as currency
	second_payment_rate = 50
dim number_of_kW as integer
dim second_payment_total as currency

'Need to find the total number of kW that qualify for this week's second payment.

second_payment_total = number_of_kW * second_payment_rate

End function