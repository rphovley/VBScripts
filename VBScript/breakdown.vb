Private Sub cmdrun_Click()

'Main set of variables
Dim rep As String
Dim rep_name As String
Dim job_id As String
Dim customer As String
Dim system_size As Double
Dim job_in_jeopardy As String
Dim cancellation As String
Dim installed As String
Dim in_progress As String
Dim sale_amount As Currency
Dim earned As Currency
Dim status As String


'row counter variable
'the naming convention for these counters follows the name of the sheet they are for
Dim financial_data_row As Integer
Dim current_data_row As Integer
Dim current_data_row2 As Integer
Dim commissions_earned_row As Integer
Dim clawbacks_pending_row As Integer
Dim jobs_in_jeopardy_row As Integer
Dim jobs_in_progress_row As Integer
Dim rep_row As Integer
'Column counters
dim financial_customer_col as integer
dim financial_repID_col as integer
dim financial_jobID_col as integer
dim financial_payment_col as integer
dim financial_kW_col as integer
	dim current_customer_col as integer
	dim current_jobID_col as integer
	dim current_systemsize_col as integer
	dim current_status_col as integer
	dim current_rep_col as integer
	dim current_jeopardy_col as integer
	dim current_cancellation_col as integer
	dim current_installed_col as integer
	dim current_inprogress_col as integer
		dim repList_rep_col as integer
		dim repList_repname_col as integer
		dim repList_repID_col as integer
		dim repList_rate_col as integer
			dim report_customer_col as integer
			dim report_jobID_col as integer
			dim report_systemsize_col as integer
			dim report_col4 as integer
			dim report_col5 as integer
			dim report_col6 as integer

set financial_customer_col = 3
set financial_jobID_col = 7
set financial_kW_col = 4
set financial_repID_col = 2
set financial_payment_col = 5
	set current_customer_col = 1
	set current_jobID_col = 2
	set current_systemsize_col = 3
	set current_status_col = 4
	set current_rep_col = 17
	set current_jeopardy_col = 19
	set current_cancellation_col = 20
	set current_installed_col = 21
	set current_inprogress_col = 22
		set repList_rate_col = 5
		set repList_rep_col = 2
		set repList_repname_col = 1
		set repList_repID_col = 3
			set report_customer_col = 1
			set report_jobID_col = 2
			set report_systemsize_col = 3
			set report_col4	= 4
			set report_col5 = 5
			set report_col6 = 6

rep_row = 2
Do
'Creates the new tabs
	call format_sheet "Commissions Earned" "Earned" "Paid" "Due"

	call format_sheet "Clawbacks Pending" "Earned" "Paid but Never Clawed Back" "Due to Evolve"
	
	call format_sheet "Jobs in Jeopardy" "Status" "Paid" "High Probability of Cancelling"
	
	call format_sheet "Jobs in Progress" "Potentially Earned" "Paid" "Potentially Due"
	
	call format_sheet "Other" "Col4" "Col5" "Col6"
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    financial_data_row = 2
    current_data_row = 2
    commissions_earned_row = 3
    clawbacks_pending_row = 3
    jobs_in_jeopardy_row = 3
    jobs_in_progress_row = 3

    with Sheets("RepList")
	    rep      = .Cells(rep_row, repList_rep_col)
	    rep_name = .Cells(rep_row, repList_repname_col)
		rep_ID   = .cells(rep_row, repList_repID_col)
		rate     = .cells(rep_row, repList_rate_col)
	end with
    
    'cycles through the accounts
    Do
        If Sheets("Current Data").Cells(current_data_row, current_rep_col) = rep Then
        
        	with Sheets("Current Data")
				'assigns values to the variables reported on
				job_id      = .Cells(current_data_row, current_jobID_col)
				customer    = .Cells(current_data_row, current_customer_col)
				system_size = .Cells(current_data_row, current_systemsize_col)
				
				job_in_jeopardy = .Cells(current_data_row, current_jeopardy_col)
				cancellation    = .Cells(current_data_row, current_cancellation_col)
				installed       = .Cells(current_data_row, current_installed_col)
				in_progress     = .Cells(current_data_row, current_inprogress_col)
			end with
			
			status = Sheets("Current Data").Cells(current_data_row, current_status_col)

			'code for calculating earned based on rep.
			with sheets("Financial Data")
				do until (IsEmpty(.cells(financial_data_row, financial_repID_col)))
					if .cells(financial_data_row, financial_repID_col) = rep_ID AND .cells(financial_data_row, financial_customer_col) = customer then
						earned = system_size * sheets("RepList").cells(rep_row, repList_rate_col)
						financial_data_row = financial_data_row + 1
					Else
						financial_data_row = financial_data_row + 1
					End if
				Loop
			end with
        
			'code for summing the amount paid for the account
			Dim sale_amount_row As Integer
			sale_amount_row = 2
			
			with sheets("Financial Data")
				Do until (IsEmpty(.cells(sale_amount_row,financial_customer_col)
					If .Cells(sale_amount_row, financial_repID_col) = rep_ID AND .cells(sale_amount_row, financial_jobID_col) = job_id Then
						sale_amount = sale_amount + .Cells(sale_amount_row, financial_payment_col)
						sale_amount_row = sale_amount_row + 1
					Else
						sale_amount_row = sale_amount_row + 1
					End If
				Loop
			end with

			'Determines which reporting sheet the account goes to
			If installed = "TRUE" Or installed = "True" Then
				with sheets("Commissions Earned")
					.Cells(commissions_earned_row, report_customer_col) = customer
					.Cells(commissions_earned_row, report_jobID_col) = job_id
					.Cells(commissions_earned_row, report_systemsize_col) = system_size
					.Cells(commissions_earned_row, report_col4) = earned
					.Cells(commissions_earned_row, report_col5) = sale_amount
					.Cells(commissions_earned_row, report_col6) = earned - sale_amount
				end with
				
				commissions_earned_row = commissions_earned_row + 1
			ElseIf cancellation = "TRUE" Or cancellation = "True" Then
				with sheets("Clawbacks Pending")
					.Cells(clawbacks_pending_row, report_customer_col) = customer
					.Cells(clawbacks_pending_row, report_jobID_col) = job_id
					.Cells(clawbacks_pending_row, report_systemsize_col) = system_size
					.Cells(clawbacks_pending_row, report_col4) = earned
					.Cells(clawbacks_pending_row, report_col5) = sale_amount
					.Cells(clawbacks_pending_row, report_col6) = earned - sale_amount
				end with
				
				clawbacks_pending_row = clawbacks_pending_row + 1
			ElseIf job_in_jeopardy = "TRUE" Or job_in_jeopardy = "True" Then
				with sheets("Jobs in Jeopardy")
					.Cells(jobs_in_jeopardy_row, report_customer_col) = customer
					.Cells(jobs_in_jeopardy_row, report_jobID_col) = job_id
					.Cells(jobs_in_jeopardy_row, report_systemsize_col) = system_size
					.Cells(jobs_in_jeopardy_row, report_col4) = status
					.Cells(jobs_in_jeopardy_row, report_col5) = sale_amount
					.Cells(jobs_in_jeopardy_row, report_col6) = 0 - sale_amount
				end with
            
				jobs_in_jeopardy_row = jobs_in_jeopardy_row + 1
			ElseIf in_progress = "TRUE" Or in_progress = "True" Then
				with sheets("Jobs in Progress")
					.Cells(jobs_in_progress_row, report_customer_col) = customer
					.Cells(jobs_in_progress_row, report_jobID_col) = job_id
					.Cells(jobs_in_progress_row, report_systemsize_col) = system_size
					.Cells(jobs_in_progress_row, report_col4) = earned
					.Cells(jobs_in_progress_row, report_col5) = sale_amount
					.Cells(jobs_in_progress_row, report_col6) = earned - sale_amount
				end with
            
				jobs_in_progress_row = jobs_in_progress_row + 1
			End If
				current_data_row = current_data_row + 1
        Else
				current_data_row = current_data_row + 1
        End If
        
    Loop Until Sheets("Current Data").Cells(current_data_row, 2) = ""
    
    
    With Worksheets("Commissions Earned").Range("A:F")
        .EntireColumn.AutoFit
    End With
    
    With Worksheets("Clawbacks Pending").Range("A:F")
        .EntireColumn.AutoFit
    End With
    
    With Worksheets("Jobs in Jeopardy").Range("A:F")
        .EntireColumn.AutoFit
    End With
    
    With Worksheets("Jobs in Progress").Range("A:F")
        .EntireColumn.AutoFit
    End With
    
    ThisWorkbook.Sheets(Array("Commissions Earned", "Clawbacks Pending", "Jobs in Jeopardy", "Jobs in Progress")).Move
    ActiveWorkbook.SaveAs ("C:\users\ezra\desktop\" & "Finance Report" & "\" & rep & ".xlsx")
    ActiveWorkbook.Close
    
    rep_row = rep_row + 1
	
Loop Until Sheets("RepList").Cells(rep_row, 2) = ""

End Sub

Sub format_sheet(ByVal sheetName As String, ByVal Col4 As String, ByVal Col5 As String, By Val Col6 As String)

Worksheets.Add(, Worksheets(Worksheets.Count)).Name = sheetName

    With Worksheets(sheetName)
        .Cells(1, 1) = sheetName
        .Cells(2, 1) = "Customer"
        .Cells(2, 2) = "JobID"
        .Cells(2, 3) = "System Size"
        .Cells(2, 4) = Col4
        .Cells(2, 5) = Col5
        .Cells(2, 6) = Col6
    End With

    'Formats the main header
    With Worksheets(sheetName).Range("A1:F1")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 77, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    'Formats the column headers
    With Worksheets(sheetName).Range("A2:F2")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With
	
End Sub

sub email(ByVal rep as string)

	'Email portion of the code
		Dim OutApp As Object
		Dim OutMail As Object

		Set OutApp = CreateObject("Outlook.Application")
		Set OutMail = OutApp.CreateItem(0)

		On Error Resume Next
	   
		With OutMail
			.To = Rep
			.CC = ""
			.BCC = ""
			.Subject = "Sales Report"
			.Body = "The attached report shows an update on your accounts."
			.Attachments.Add ("C:\users\ezra\desktop\" & "Finance Report" & "\" & rep & ".xlsx")
			'.Display
			.Send
		End With
		On Error GoTo 0

		Set OutMail = Nothing
		Set OutApp = Nothing
		
	Loop Until Sheets("Reps").Cells(reprow, repcol) = ""

	End If
	
End Sub
