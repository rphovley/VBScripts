function first_payment()

''Occurs when docs signed = "Y"
dim number_of_accounts as integer
dim one_two_payment as currency
	one_two_payment = 250
dim three_five_payment as currency
	three_five_payment = 350
dim six_plus_payment as currency
	six_plus_payment = 450
dim first_payment_total as currency

'Needs to first count how many accounts qualify for this week's  first payment

	if number_of_accounts <= 2 And number_of_accounts > 0 then
		first_payment_total = number_of_accounts * one_two_payment
	ElseIf number_of_accounts > 2 and number_of_accounts <= 5 then
		first_payment_total = number_of_accounts * three_five_payment
	ElseIf number_of_accounts > 5 then
		first_payment_total = number_of_accounts * six_plus_payment
	End If

	return first_payment_total
	
End function