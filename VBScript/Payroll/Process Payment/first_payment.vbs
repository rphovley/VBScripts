function first_payment(byref currentRep as cRepData, byref SalesThisWeek as Integer)

''Occurs when site survey complete or 2 weeks after job created date
dim new_one_two as currency
	new_one_two = 100
dim new_three_five as currency
	new_three_five = 200
dim new_six_plus as currency
	new_six_plus = 300
	
''Occurs when final contract is signed
dim number_of_accounts as integer
dim old_one_two as currency
	old_one_two = 250
dim old_three_five as currency
	old_three_five = 350
dim old_six_plus as currency
	old_six_plus = 450
dim first_payment_total as currency

'Needs to first count how many accounts qualify for this week's  first payment
 Dim customerCol, jobCol, kWCol, statusCol, subStatusCol, _
        createdDateCol, repEmailCol, isDocSignedCol, isFinalContractCol As Integer

        customerCol = 1
        jobCol = 2
        createdDateCol = 6
        repEmailCol = 9
        
    ''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)

    ''''''''''''''''''''''''''Create Nate's Evolution'''''''''''''
    	createNatesEvo(workBookName)
    	
    ''''''''''''''''''''''''''''''Worksheets''''''''''''''''''''''
    Dim jobDataSheet, printSheet As Worksheet
        Set jobDataSheet = NatesEvolution.Worksheets("Nate's Evolution")
        Set printSheet = NatesEvolution.Worksheets("Debug")

    ''''''''''''''''''''''''''''''Row Counters''''''''''''''''''''''
    Dim inputRow, printRow As Integer
        printRow = 2

    '''''''''''''''''''''''''''''Input Object''''''''''''''''''''''
    Dim jobData() As cJobData
    Dim currentJob  As cJobData

    '''''''''''''''''''''''''''''Data Size'''''''''''''''''''''''''
    Dim jobDataSize As Long
    	jobDataSize = jobDataSheet.Cells(1,1).End(xlDown).Row - 2
    	ReDim jobData(jobDataSize)

''''''''''''''''''''''''''''''Calculates the first payment''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''FOR ACCOUNTS LESS THAN 60 DAYS''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if number_of_accounts <= 2 And number_of_accounts > 0 then
		first_payment_total = number_of_accounts * new_one_two
	ElseIf number_of_accounts > 2 and number_of_accounts <= 5 then
		first_payment_total = number_of_accounts * new_three_five
	ElseIf number_of_accounts > 5 then
		first_payment_total = number_of_accounts * new_six_plus
	End If


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''FOR ACCOUNTS GREATER THAN 60 DAYS''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	if number_of_accounts <= 2 And number_of_accounts > 0 then
		first_payment_total = number_of_accounts * old_one_two
	ElseIf number_of_accounts > 2 and number_of_accounts <= 5 then
		first_payment_total = number_of_accounts * old_three_five
	ElseIf number_of_accounts > 5 then
		first_payment_total = number_of_accounts * old_six_plus
	End If

	return first_payment_total
	
End function