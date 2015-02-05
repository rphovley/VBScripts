	'Columns for the "Report Tab"'
	Dim repJobIdCol, repDateCol, repkWCol, repStatusCol, _
	 repOldNewCol, repPaidOutCol, repCurValCol, repEstCol, _
	 repActCol, repCheckCol, repPermitCol As Integer

	'Columns for the "Master Report" Tab'
	Dim masJobIdCol, masDateCol, maskWCol, masStatusCol, _
	 masFinalCol, masInstallCol, masInstallDateCol, masCancelledCol, masPermitCol As Integer


	 'Collection KEYS'
	Dim dJOBID, dKW, dSTATUS, dPERMITSTATUS, dALREADYPAID, dDATE, dFINAL, dINSTALL, dINSTALLDATE, dCANCELLED AS String
	
	Dim full_value as currency
	dim booster as currency
	dim cancel_value as currency
	Dim kW as double

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''Main Sub for Estimate'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Main sub for creating estimate'
'Works with the "SolarCity Audit.xlsm" file in the Historical Breakdowns folder'
Sub createEstimate()

	'initialize variables from above'
	initVar

	'dim the row vars for both tabs'
	Dim MasterReportRow, ReportRow, boostRow As Integer
	Dim alreadyPaid As Double
	Dim commishWorkbook As String

	MasterReportRow = 2
	ReportRow       = 2
	alreadyPaid     = 0
 
 	'get workbook name of the commission workbook'
    'commishWorkbook = convertToName(Application.GetOpenFilename())

    'boostRow = InputBox("What is the starting row in the 'Accounting Summary' tab for the boost payments?", "Boost Payments Row")
	'used to pass information back and forth from functions'
	Dim dataFromMasterReport As New Collection

	'Loop through the Master Report'
	Do Until isEmpty(Sheets("Master Report").Cells(MasterReportRow, 1).Value)
	
		'Collect Data from Master Report and Determine what should be paid out to us in the Master Report'
		Set dataFromMasterReport = determinePayout(dataFromMasterReport, MasterReportRow)
		
		Call check_structure(dataFromMasterReport, ReportRow, repDateCol, repOldNewCol, repkWCol, MasterReportRow, masCancelledCol)
		
		'set what was paid out in the commissions sheet'
		'alreadyPaid = whatWasPaid(dataFromMasterReport.Item(dJOBID), commishWorkbook, boostRow)
		dataFromMasterReport.Add alreadyPaid, dALREADYPAID
		
		'print out what should be paid out in the Report Tab'
	 	printData dataFromMasterReport, ReportRow

	 	'In order to reset the values in a collection the values have to be removed first, this function does that'
		Set dataFromMasterReport = refreshCollection(dataFromMasterReport)
		dataFromMasterReport.Remove dALREADYPAID

	 	MasterReportRow = MasterReportRow + 1
	 	ReportRow       = ReportRow + 1
	Loop


End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''Supporting Subs and Functions''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Determine what should be paid out to us in the Master Report'
Function determinePayout(ByRef dataFromMasterReport As Collection, ByVal MasterReportRow As Integer) As Collection

	'Collection data from Master Report'
	Set	dataFromMasterReport = setCollection(dataFromMasterReport, MasterReportRow)

	'We make some decision based on what we find in the report'

	'This is returning the collection to the calling sub'
	
	Set determinePayout = dataFromMasterReport
End Function

'Set Collection Values for the data from the Master Report'
Function setCollection(ByRef dataFromMasterReport As Collection, ByVal MasterReportRow As Integer) As Collection

	With Sheets("Master Report")
		dataFromMasterReport.Add .Cells(MasterReportRow, masJobIdCol), dJOBID
	    dataFromMasterReport.Add .Cells(MasterReportRow, maskWCol), dKW
	    dataFromMasterReport.Add .Cells(MasterReportRow, masStatusCol), dSTATUS
		dataFromMasterReport.Add .Cells(MasterReportRow, masPermitCol), dPERMITSTATUS
	    dataFromMasterReport.Add .Cells(MasterReportRow, masDateCol), dDATE
	    dataFromMasterReport.Add .Cells(MasterReportRow, masFinalCol), dFINAL
	    dataFromMasterReport.Add .Cells(MasterReportRow, masCancelledCol), dCANCELLED
	    dataFromMasterReport.Add .Cells(MasterReportRow, masInstallDateCol), dINSTALLDATE
	    dataFromMasterReport.Add .Cells(MasterReportRow, masInstallCol), dINSTALL
	End With

	'this is returning the collection from the calling function'
	Set setCollection = dataFromMasterReport

End Function

'In order to reset the values in a collection the values have to be removed first, this function does that'
Function refreshCollection(ByRef dataFromMasterReport As Collection) As Collection
	
	dataFromMasterReport.Remove dJOBID
    dataFromMasterReport.Remove dKW
    dataFromMasterReport.Remove dSTATUS
	dataFromMasterReport.Remove dPERMITSTATUS
    dataFromMasterReport.Remove dDATE
    dataFromMasterReport.Remove dFINAL
    dataFromMasterReport.Remove dCANCELLED
    dataFromMasterReport.Remove dINSTALLDATE
    dataFromMasterReport.Remove dINSTALL

	Set refreshCollection = dataFromMasterReport
End Function
'Sub to print out data gathered into the Report Tab'
Sub printData(ByRef dataFromMasterReport, ByVal ReportRow As Integer)
	
	With Sheets("Report")
		.Cells(ReportRow, repJobIdCol).Value   = dataFromMasterReport.Item(dJOBID)
		.Cells(ReportRow, repDateCol).Value    = dataFromMasterReport.Item(dDATE)
		.Cells(ReportRow, repkWCol).Value      = dataFromMasterReport.Item(dKW)
		.Cells(ReportRow, repStatusCol).Value  = dataFromMasterReport.Item(dSTATUS)
		.Cells(ReportRow, repPermitCol).Value  = dataFromMasterReport.Item(dPERMITSTATUS)
		.Cells(ReportRow, repPaidOutCol).Value = dataFromMasterReport.Item(dALREADYPAID)
		' .Cells(ReportRow, repOldNewCol).Value = dataFromMasterReport.Item(dOLDNEw)
		' .Cells(ReportRow, repEstCol).Value    = dataFromMasterReport.Item(dEST)
		' .Cells(ReportRow, repActCol).Value    = dataFromMasterReport.Item(dACT)
		' .Cells(ReportRow, repCheckCol).Value  = dataFromMasterReport.Item(dCheck)
		
	End With

End Sub

'initialize variables for columns'
Sub initVar()

     'Columns for the "Report Tab"'
	 repJobIdCol   = 1
	 repDateCol    = 2
	 repkWCol      = 3
	 repStatusCol  = 4
	 repPermitCol  = 5
	 repOldNewCol  = 6
	 repPaidOutCol = 7
	 repCurValCol  = 8
	 repEstCol     = 9
	 repActCol     = 10
	 repCheckCol   = 11


	 'Columns for the "Master Report" Tab'
	 masJobIdCol       = 2
	 masDateCol        = 7
	 maskWCol          = 3
	 masStatusCol      = 4
	 masPermitCol      = 5
	 masFinalCol       = 8
	 masInstallDateCol = 9
	 masCancelledCol   = 20
	 masInstallCol     = 21

	 'Collection Keys'
	 dJOBID       = "jobID"
	 dKW          = "kW"
	 dSTATUS      = "Status"
	 dPERMITSTATUS= "PermitStatus"
	 dALREADYPAID = "AlreadyPaid"
	 dDATE        = "Date"
	 dFINAL       = "Final"
	 dINSTALL     = "Installed"
	 dINSTALLDATE = "dateInstalled"
	 dCANCELLED   = "Cancelled"

End Sub

'Checks which payout structure this account falls under
Sub check_structure(ByRef dataFromMasterReport As Collection, ByVal ReportRow, ByVal repDateCol, ByVal repOldNewCol,ByVal kWCol, ByVal MasterReportRow, ByVal masCancelledCol)
    With Sheets("Report")
		kW = .cells(ReportRow, repkWCol).Value
        If .Cells(ReportRow, repDateCol) < 41974 Then
            .Cells(ReportRow, repOldNewCol) = "Old"

            Call old_payout_structure(dataFromMasterReport, reportRow)

        Else
            .Cells(ReportRow, repOldNewCol) = "New"
            Call new_payout_structure (MasterReportRow, masFinalCol, masInstallCol, ReportRow, repCurValCol, kW, masCancelledCol)
        End If
    End With
End Sub

'Sub for New Payout Structure
Sub new_payout_structure(ByVal MasterReportRow, ByVal masFinalCol, ByVal masInstallCol, ByVal ReportRow, ByVal repCurValCol, ByVal kW, ByVal masCancelledCol)
		full_value = kW * 500 * 1.0
		booster = kW * 500 * .5
		cancel_value = 0
	With Sheets("Master Report")
		If .cells(MasterReportRow, masCancelledCol) <> "" then
			Sheets("Report").cells(ReportRow, repCurValCol) = cancel_value
		Else
			If Sheets("Report").cells(ReportRow, repStatusCol) = "Cancelled" Or Sheets("Report").cells(ReportRow, repStatusCol) = "Sales" then
				If Sheets("Report").cells(ReportRow, repPermitCol) = "Account Cancelled" , "Customer Uncertain" , "Customer Unresponsive" , "On Hold" , "Pending NOC" , "Pending Save" then
					Sheets("Report").cells(ReportRow, repCurValCol) = - Sheets("Report").cells(ReportRow, repPaidOutCol)
				End If			
			Else
				If .cells(MasterReportRow, masFinalCol) <> "" And .cells(MasterReportRow, masInstallCol) <> "" then
					Sheets("Report").cells(ReportRow, repCurValCol) = full_value
				ElseIf .cells(MasterReportRow, masFinalCol) <> "" And .cells(MasterReportRow, masInstallCol) = "" then
					Sheets("Report").cells(ReportRow, repCurValCol) = booster	
				Else
					Sheets("Report").cells(ReportRow, repCurValCol) = cancel_value
				End If
			End If
		End If
	End With

End Sub

'Sub for Old payout structure
Sub old_payout_structure(ByRef dataFromMasterReport As Collection, ByVal ReportRow As Integer)
	Dim paymentAmount As Double
	Dim todaysDate As Date
	Dim diffClosed, diffInstall, dateCreated As Integer
	Const MIN_DATE = #2/28/2014#
	Const MAX_DATE = #5/1/2014#
	todaysDate = Date()


	'Days between created date (closed won) and todays date'
	dateCreated = dataFromMasterReport.Item(dDATE)
	diffClosed = DateDiff("d", todaysDate, dateCreated)

	Dim test As Date 
	test = dataFromMasterReport.Item(dINSTALLDATE)

	'Days between install date and todays date'
	If dataFromMasterReport.Item(dINSTALLDATE) <> "12:00am" Then
		diffInstall = DateDiff("d", todaysDate, dataFromMasterReport.Item(dINSTALLDATE))
	End If

	'Should this be a final payment die to install?'
	If diffInstall > 30 Then
		'Should be at final payment'
		'checks to see if it should be paid out at $600 per kW'
		If dateCreated > MIN_DATE AND dateCreated < MAX_DATE Then
			paymentAmount = dataFromMasterReport.Item(dKW) * 600
		Else
			paymentAmount = dataFromMasterReport.Item(dKW) * 500
		End If
	Else
		If diffClosed > 180 Then
			'Should be at final payment'
			'checks to see if it should be paid out at $600 per kW'
			If dateCreated > MIN_DATE AND dateCreated < MAX_DATE Then
				paymentAmount = dataFromMasterReport.Item(dKW) * 600
			Else
				paymentAmount = dataFromMasterReport.Item(dKW) * 500
			End If
		ElseIf diffClosed > 90 Then
			'Should be at 2nd payment'
			'checks to see if it should be paid out at $600 per kW'
			If dateCreated > MIN_DATE AND dateCreated < MAX_DATE Then
				paymentAmount = dataFromMasterReport.Item(dKW) * 600 * .75
			Else
				paymentAmount = dataFromMasterReport.Item(dKW) * 500 * .75
			End If
		ElseIF diffClosed > 30 Then
			'Should be at 1st payment'
			'checks to see if it should be paid out at $600 per kW'
			If dateCreated > MIN_DATE AND dateCreated < MAX_DATE Then
				paymentAmount = dataFromMasterReport.Item(dKW) * 600 * .5
			Else
				paymentAmount = dataFromMasterReport.Item(dKW) * 500 * .5
			End If
		End If

		Sheets("Report").cells(ReportRow, repCurValCol) = paymentAmount

	End If


	

End Sub
'Function to return the amount that was paid out by Solar City for a specific JobID'
Function whatWasPaid(ByVal jobID As String, ByVal Workbook As String, ByVal boostRow As Integer) As Double
	'variables for tab names'
	Dim firsts, seconds, finals, pos, neg, acc As String
		firsts  = "1st_Payment"
		seconds = "2nd_Payment"
		finals  = "Final_Payment"
		pos     = "Pos_Payment"
		neg     = "Neg_Payment"
		acc     = "Accounting Summary"

	'Columns First Payment Tab'
	Dim firstDateCol, firstJobIDCol, firstKWCol, firstPaymentCol, firstPayDateCol As Integer
		firstDateCol    = 3
		firstJobIDCol   = 4
		firstKWCol      = 6
		firstPaymentCol = 21
		firstPayDateCol = 22

	'Columns Other Payment Tabs'
	Dim othDateCol, othJobIDCol, othKWCol, othPaymentCol, othPayDateCol As Integer
		othDateCol    = 3
		othJobIDCol   = 4
		othKWCol      = 7
		othPaymentCol = 22
		othPayDateCol = 23

	'Boost Payment Columns'
	Dim booJobIDCol, booPaymentCol As Integer
		booJobIDCol   = 2
		booPaymentCol = 7

	'Row counter'
	Dim currentRow As Integer
	'value that will be returned'
	Dim paymentAmount As Double
	paymentAmount = 0
	

		'loops through each tab and gets the payment amount for the related jobID'
		paymentAmount = tabLoop(Workbook, firsts, jobID, firstJobIDCol, firstPaymentCol, paymentAmount)
		paymentAmount = tabLoop(Workbook, seconds, jobID, othJobIDCol, othPaymentCol, paymentAmount)
		paymentAmount = tabLoop(Workbook, finals, jobID, othJobIDCol, othPaymentCol, paymentAmount)
		paymentAmount = tabLoop(Workbook, pos, jobID, othJobIDCol, othPaymentCol, paymentAmount)
		paymentAmount = tabLoop(Workbook, neg, jobID, othJobIDCol, othPaymentCol, paymentAmount)
		paymentAmount = tabLoop(Workbook, acc, jobID, booJobIDCol, booPaymentCol, paymentAmount)

	whatWasPaid = paymentAmount

End Function

'loops through each tab and gets the payment amount for the related jobID'
Function tabLoop(ByVal Workbook As String, ByVal SheetName As String,  jobID As String, ByVal jobIDCol As Integer, byVal paymentCol As Integer, byVal paymentAmount As Double) As Double
	'Go through the Payment Tab and find any relevant payments'
	Dim currentRow As Integer
	currentRow = 1

	With Workbooks(WorkbooK).Sheets(SheetName)
		Do Until isEmpty(.Cells(currentRow, 1))
			If .Cells(currentRow, jobIDCol).Value = jobID Then
				paymentAmount = paymentAmount + .Cells(currentRow, paymentCol)
			End IF
			currentRow = currentRow + 1
		Loop
	End With

	tabLoop = paymentAmount
End Function

'function to convert a filepath to a fileName'
Function convertToName(ByVal Path As String) As String

     For Each wbk1 In Workbooks
        If (wbk1.Path & "\" & wbk1.Name = Path) Then
            convertToName = wbk1.Name
            Exit For
        End If
    Next
End Function

Sub estimated_payment(ByVal ReportRow, ByVal repPaidOutCol, byVal repCurValCol, byVal repEstCol)
	
	With Sheets("Report")
		.cells(ReportRow, repEstCol) = .cells(ReportRow, repCurValCol) - .cells(ReportRow, repPaidOutCol)
	End With
	
End Sub

Sub check_payments(ByVal ReportRow, ByVal repEstCol, byVal repActCol, byVal repCheckCol)

	With Sheets("Report")
		If .cells(ReportRow, repEstCol) = .cells(ReportRow, repActCol) then
			.cells(ReportRow, repCheckCol) = "TRUE"
		Else
			.cells(ReportRow, repCheckCol) = "FALSE"
			.cells(ReportRow, repCheckCol).Interior.color = RGB(255,0,0)
		End If
	End With

End Sub

'is status a cancelled status'
Function isJobCancelled(ByRef Status As String) As Boolean

    Dim isArray As Variant
    isArray = Array("Customer Uncertain", "Customer Unresponsive", _
        "Job Disqualified", "On Hold", "Cancelled")
    
    For Each permitStatus In isArray
    
        If permitStatus = Status Then
            isJobCancelled = True
            Exit For
        End If
    Next permitStatus

