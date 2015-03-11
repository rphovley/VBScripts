function first_payment(byref currentRep as cRepData, byref SalesThisWeek as Integer)

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
	if number_of_accounts <= 2 And number_of_accounts > 0 then
		first_payment_total = number_of_accounts * one_two_payment
	ElseIf number_of_accounts > 2 and number_of_accounts <= 5 then
		first_payment_total = number_of_accounts * three_five_payment
	ElseIf number_of_accounts > 5 then
		first_payment_total = number_of_accounts * six_plus_payment
	End If

	return first_payment_total
	
End function