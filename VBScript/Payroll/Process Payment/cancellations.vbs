Function cancells(ByRef currentJob As cJobData, ByRef currentRep as cRepData, ByVal WorkBookName As String) As cJobData
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'''COLUMNS'''''
Dim jobCol, emailCol, kWCol, paidCol, clawedCol, dateCol As Integer
	
	jobCol    = 1
	emailCol  = 2
	kWCol     = 3
	paidCol   = 4
	clawedCol = 5
	dateCol   = 6

Dim row As Integer
	row = 2
	

Dim isAlreadyCancelled As Boolean
Dim whatWasPaid As Currency

'Check if job is now cancelled'
	If currentJob.IsCancelled Then

		With Workbooks(WorkBookName).WorkSheets("Cancelled")
			Do Until IsEmpty(.Cells(row, 1).value)
			'''''''''''''''''''''''''''''''QUESTION:::::IF THERE ISN'T ENOUGH INSTALL MONEY TO COVER THE JOB, DO WE TAKE OUT PART?''''''''''''''''''''''
				'If it is on the list and can't be taken out of installs, take out as much as we can from install Pool"
				If currentJob.JobID = .Cells(row, jobCol).value AND .Cells(row, paidCol).value > currentRep.InstallPool Then				
					'Amount of Job Cancellation'
					
					whatWasPaid = .Cells(row, paidCol).value

					'Take out as much as we can from Install Pool'
					currentJob.ClawbackAmount = whatWasPaid - currentRep.InstallPool

					'Make adjustment in record to reflect what has been clawed back'
					.Cells(row, paidCol).value = .Cells(row, paidCol).value - currentJob.ClawbackAmount

					isAlreadyCancelled = True
					Exit Do

				'If it is on the list and it can be taken out of installs completely, take it out of installs and remove from list'
				ElseIf currentJob.JobID = .Cells(row, jobCol).value AND .Cells(row, paidCol).value < currentRep.InstallPool Then
					'took out of install pool'
					currentRep.InstallPool = currentRep.InstallPool - .Cells(row, paidCol).value
					'set the clawback amount'
					currentJob.ClawbackAmount = .Cells(row, paidCol).value
					'Remove from list'
					Rows(row).EntireRow.Delete

					isAlreadyCancelled = True
					Exit Do
				End If
				row = row + 1
			Loop
			'If it is not on the list and cannot be taken out of installs. Take out what we can and print remaining amount it to the list'
			If NOT isAlreadyCancelled AND currentJob.WhatWasPaid > currentRep.InstallPool Then
				'Amount of Job Cancellation'
					whatWasPaid = currentJob.WhatWasPaid

					'Take out as much as we can from Install Pool'
					currentJob.ClawbackAmount = currentRep.InstallPool

					'Make adjustment in record to reflect what has been clawed back'
					currentJob.WhatWasPaid = whatWasPaid - currentJob.ClawbackAmount
					printCancellation currentJob, workBookName 

			'If it is not on the list and it can be taken out of installs completely, take it out of installs and don't add to list'
			ElseIf NOT isAlreadyCancelled AND currentJob.WhatWasPaid < currentRep.InstallPool Then
					whatWasPaid = currentJob.WhatWasPaid

				'Set clawback amount to the amount that we have paid out'
				currentJob.ClawbackAmount = whatWasPaid
				currentJob.WhatWasPaid = 0
				'Remove clawback amount from install Pool'
				currentRep.InstallPool = currentRep.InstallPool - currentJob.ClawbackAmount

			End If
		End With


		'ALL CASES SHOULD REMOVE THIS JOB FROM FIRST AND SECOND PAYMENTS TABS'
	End If
	payroll_main.repData.Remove currentRep.Email
	payroll_main.repData.Add currentRep.Email, currentRep

	Set cancells = currentJob
End Function

Sub printCancellation(ByRef currentJob As cJobData, ByVal WorkBookName As String)
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''


'''COLUMNS'''''
Dim jobCol, emailCol, kWCol, paidCol, clawedCol, dateCol As Integer
	
	jobCol    = 1
	emailCol  = 2
	kWCol     = 3
	paidCol   = 4
	clawedCol = 5
	dateCol   = 6
        
''''''''''''''''''''''''''''''Workbooks''''''''''''''''''''''
        'workBookName = InputBox("What is the master report's name?") & ".xlsx"
    Dim NatesEvolution As Workbook
        Set NatesEvolution = Workbooks(workBookName)
	
	'worksheets'
	Dim Cancelled As Worksheet
		Set Cancelled = NatesEvolution.Worksheets("Cancelled")
	'Row'
	Dim printRow As Integer
		printRow = Cancelled.UsedRange.Rows.Count + 1

	With Cancelled
		.Cells(printRow, jobCol)    = currentJob.JobID
		.Cells(printRow, emailCol)  = currentJob.repEmail
		.Cells(printRow, kWCol)     = currentJob.kW
		.Cells(printRow, paidCol)   = currentJob.WhatWasPaid
		.Cells(printRow, clawedCol) = currentJob.ClawbackAmount
		.Cells(printRow, dateCol)   = Date()
	End With

End Sub