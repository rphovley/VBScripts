'psuedo code for payment system'

'first_payment() *happens at docs signed'
	'- needs a count of accounts for that week'
	'- using the pay scale, determine rate to advance'

'second_payment() *happens at final contract'
	'-receives $50/kW'

'final_payment() *happens at Install'
	'-amount_for_install = total_due - what_was_paid'
	'*total_due is determined by the following'
		'If they are a "normal" rep, they are paid using the pay scales'
		'If they are an "experienced" rep (more than 300kW) they are on a sliding pay scale'
		'We need to flesh out Marketing Events pay structure
		'There may be other issues to find out here'

	''''****Cancellations are accounted for here*****'''''
	'to_be_paid = amount_for_install - cancellation_pool'

	'cancellation_pool = cancellation_pool - amount_for_install'   
		'*This is to adjust the cancellation_pool to reflect installs taken from it'

	'remove job from "pending" tab'

'clawbacks() *any pending sale that is now cancelled'

	'cancellation_pool = cancellation_pool - what_was_paid'

	'post cancellation pool data to sheet'
	
	'remove job from "pending" tab'

'update_slider() *updates reps who should be on a sliding pay scale'
	'loop through all sales in "Nate's Evolution" and sum up kW for each rep'
	'if the rep is more than 300kW he is paid on a sliding scale'
	'*note, the job that puts the rep over the 300kW is considered the first to be paid on sliding scale'

'Implementaion Notes'
	
	'BEFORE PROCESS'
		'Sort Cancelled sheet WhatWasPaid from smallest to largest'
		'Sort Master Spreadsheet created Date from oldest to newest'
		'Check Nate's Evolution that the formulas include all the data from Master and Doc Signed Input'
		'Remove any job duplicates from MasterReport'

	'ADJUSTMENTS AFTER PROCESS HAS BEEN RUN'
		'Bobby and Josh will press a button, and based on the normal conditions, they will'
		'receive a printed list back in the "Hunt's" tab that includes all the reps to be paid'
		'out, as well as any that were not paid out because of negatives.  If a particular rep'
		'is to receive money in spite of a cancellation, the amount he is to receive should'
		'be added as an "advance" transaction instead of making any edits to the amount clawed back'
		'This will make for cleaner code and it will be easier to see when an exception is happening'
		'when we look through the data'

'