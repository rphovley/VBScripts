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
Dim system_value As Currency
Dim status As String
Dim rate As Currency


'row counter variable
'the naming convention for these counters follows the name of the sheet they are for
Dim current_data_row As Integer
Dim current_data_row2 As Integer
Dim commissions_earned_row As Integer
Dim clawbacks_pending_row As Integer
Dim jobs_in_jeopardy_row As Integer
Dim jobs_in_progress_row As Integer
Dim rep_row As Integer
Dim other_row As Integer
Dim financial_row As Integer
'Column counters
Dim financial_customer_col As Integer
Dim financial_repID_col As Integer
Dim financial_jobID_col As Integer
Dim financial_payment_col As Integer
Dim financial_kW_col As Integer
    Dim current_customer_col As Integer
    Dim current_jobID_col As Integer
    Dim current_systemsize_col As Integer
    Dim current_status_col As Integer
    Dim current_rep_col As Integer
    Dim current_jeopardy_col As Integer
    Dim current_cancellation_col As Integer
    Dim current_installed_col As Integer
    Dim current_inprogress_col As Integer
        Dim repList_rep_col As Integer
        Dim repList_repname_col As Integer
        Dim repList_repID_col As Integer
        Dim repList_rate_col As Integer
            Dim report_customer_col As Integer
            Dim report_jobID_col As Integer
            Dim report_systemsize_col As Integer
            Dim report_amount As Integer
            Dim report_col5 As Integer
            Dim report_col6 As Integer

financial_customer_col = 3
financial_jobID_col = 7
financial_kW_col = 4
financial_repID_col = 2
financial_payment_col = 5
    current_customer_col = 1
    current_jobID_col = 2
    current_systemsize_col = 3
    current_status_col = 4
    current_rep_col = 17
    current_jeopardy_col = 19
    current_cancellation_col = 20
    current_installed_col = 21
    current_inprogress_col = 22
        repList_rate_col = 5
        repList_rep_col = 2
        repList_repname_col = 1
        repList_repID_col = 3
            report_customer_col = 1
            report_jobID_col = 2
            report_systemsize_col = 3
            report_col4 = 4
            report_col5 = 5
            report_col6 = 6

rep_row = 2
Do
'Creates the new tabs
    Call format_sheet("Jobs Installed", "System Value", "Paid", "Due")

    Call format_sheet("Clawbacks Pending", "System Value", "Paid but Never Clawed Back", "Due to Evolve")
    
    Call format_sheet("Jobs in Jeopardy", "Status", "Paid", "High Probability of Cancelling")
    
    Call format_sheet("Jobs in Progress", "System Value", "Paid", "Potentially Due")
    
    Call format_sheet("Other", "System Value", "Amount", "")
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    current_data_row = 2
    commissions_earned_row = 3
    clawbacks_pending_row = 3
    jobs_in_jeopardy_row = 3
    jobs_in_progress_row = 3
    other_row = 3

    With Sheets("RepList")
        rep = .Cells(rep_row, repList_rep_col)
        rep_name = .Cells(rep_row, repList_repname_col)
        rep_ID = .Cells(rep_row, repList_repID_col)
        rate = .Cells(rep_row, repList_rate_col)
    End With
    
    'cycles through the accounts for the first four tabs
    Do
        If Sheets("Current Data").Cells(current_data_row, current_rep_col) = rep Then
        
            With Sheets("Current Data")
                'assigns values to the variables reported on
                job_id = .Cells(current_data_row, current_jobID_col)
                customer = .Cells(current_data_row, current_customer_col)
                system_size = .Cells(current_data_row, current_systemsize_col)
                
                job_in_jeopardy = .Cells(current_data_row, current_jeopardy_col)
                cancellation = .Cells(current_data_row, current_cancellation_col)
                installed = .Cells(current_data_row, current_installed_col)
                in_progress = .Cells(current_data_row, current_inprogress_col)
            End With
            
            status = Sheets("Current Data").Cells(current_data_row, current_status_col)

            'code for calculating earned based on rep.
            system_value = system_size * rate

        
            'code for summing the amount paid for the account
            Dim sale_amount_row As Integer
            sale_amount_row = 2
            sale_amount = 0
            With Sheets("Financial Data")
                Do Until (IsEmpty(.Cells(sale_amount_row, financial_repID_col)))
                    If .Cells(sale_amount_row, financial_repID_col) = rep_ID And .Cells(sale_amount_row, financial_jobID_col) = job_id Then
                        sale_amount = sale_amount + .Cells(sale_amount_row, financial_payment_col)
                        sale_amount_row = sale_amount_row + 1
                    Else
                        sale_amount_row = sale_amount_row + 1
                    End If
                Loop
            End With

            'Determines which reporting sheet the account goes to
            If installed = "TRUE" Or installed = "True" Then
                With Sheets("Jobs Installed")
                    .Cells(commissions_earned_row, report_customer_col) = customer
                    .Cells(commissions_earned_row, report_jobID_col) = job_id
                    .Cells(commissions_earned_row, report_systemsize_col) = system_size
                    .Cells(commissions_earned_row, report_col4) = system_value
                    .Cells(commissions_earned_row, report_col5) = sale_amount
                    .Cells(commissions_earned_row, report_col6) = system_value - sale_amount
                End With
                
                commissions_earned_row = commissions_earned_row + 1
            ElseIf cancellation = "TRUE" Or cancellation = "True" Then
                With Sheets("Clawbacks Pending")
                    .Cells(clawbacks_pending_row, report_customer_col) = customer
                    .Cells(clawbacks_pending_row, report_jobID_col) = job_id
                    .Cells(clawbacks_pending_row, report_systemsize_col) = system_size
                    .Cells(clawbacks_pending_row, report_col4) = system_value
                    .Cells(clawbacks_pending_row, report_col5) = sale_amount
                    .Cells(clawbacks_pending_row, report_col6) = sale_amount
                End With
                
                clawbacks_pending_row = clawbacks_pending_row + 1
            ElseIf job_in_jeopardy = "TRUE" Or job_in_jeopardy = "True" Then
                With Sheets("Jobs in Jeopardy")
                    .Cells(jobs_in_jeopardy_row, report_customer_col) = customer
                    .Cells(jobs_in_jeopardy_row, report_jobID_col) = job_id
                    .Cells(jobs_in_jeopardy_row, report_systemsize_col) = system_size
                    .Cells(jobs_in_jeopardy_row, report_col4) = status
                    .Cells(jobs_in_jeopardy_row, report_col5) = sale_amount
                    .Cells(jobs_in_jeopardy_row, report_col6) = 0 - sale_amount
                End With
            
                jobs_in_jeopardy_row = jobs_in_jeopardy_row + 1
            ElseIf in_progress = "TRUE" Or in_progress = "True" Then
                With Sheets("Jobs in Progress")
                    .Cells(jobs_in_progress_row, report_customer_col) = customer
                    .Cells(jobs_in_progress_row, report_jobID_col) = job_id
                    .Cells(jobs_in_progress_row, report_systemsize_col) = system_size
                    .Cells(jobs_in_progress_row, report_col4) = system_value
                    .Cells(jobs_in_progress_row, report_col5) = sale_amount
                    .Cells(jobs_in_progress_row, report_col6) = system_value - sale_amount
                End With
            
                jobs_in_progress_row = jobs_in_progress_row + 1
            End If
                current_data_row = current_data_row + 1
        Else
                current_data_row = current_data_row + 1
        End If
        
    Loop Until Sheets("Current Data").Cells(current_data_row, 2) = ""
    
    'Loop for the "Other" sheet
    
    financial_row = 2
    With Sheets("Financial Data")
        Do Until (IsEmpty(.Cells(financial_row, financial_repID_col)))
            If .Cells(financial_row, financial_repID_col) = rep_ID Then
                If .Cells(financial_row, financial_jobID_col) = "Sunnova" Or .Cells(financial_row, financial_jobID_col) = "" Then
                    customer = .Cells(financial_row, financial_customer_col)
                    job_id = .Cells(financial_row, financial_jobID_col)
                    system_size = .Cells(financial_row, financial_kW_col)
                    system_value = system_size * rate
                    sale_amount = .Cells(financial_row, financial_payment_col)
                    
                    With Sheets("Other")
                        .Cells(other_row, report_customer_col) = customer
                        .Cells(other_row, report_jobID_col) = job_id
                        .Cells(other_row, report_systemsize_col) = system_size
                        .Cells(other_row, report_col4) = system_value
                        .Cells(other_row, report_col5) = sale_amount
                    End With
                    
                    other_row = other_row + 1
                End If
                
                financial_row = financial_row + 1
            Else
                financial_row = financial_row + 1
            End If
        Loop
    End With
    
    With Worksheets("Jobs Installed")
		.cells(commissions_earned_row, report_systemsize_col) = "Total:"
		.cells(commissions_earned_row, report_col4).Formula = "=Sum(" & Range(cells(3, report_col4), cells(commissions_earned_row - 1, report_col4)).Address() & ")"
		.cells(commissions_earned_row, report_col5).Formula = "=Sum(" & Range(cells(3, report_col5), cells(commissions_earned_row - 1, report_col5)).Address() & ")"
		.cells(commissions_earned_row, report_col6).Formula = "=Sum(" & Range(cells(3, report_col6), cells(commissions_earned_row - 1, report_col6)).Address() & ")"

        .Range("A:F").EntireColumn.AutoFit
		For x = 3 to commissions_earned_row
			if .cells(x, report_col4) < 0 then
				.cells(x, report_col4).Font.ColorIndex = 3
			End if
			if .cells(x, report_col5) < 0 then
				.cells(x, report_col5).Font.ColorIndex = 3
			End if
			if .cells(x, report_col6) < 0 then
				.cells(x, report_col6).Font.ColorIndex = 3
			End if
		Next x
    End With
    
    With Worksheets("Clawbacks Pending")
		.cells(clawbacks_pending_row, report_systemsize_col) = "Total:"
		.cells(clawbacks_pending_row, report_col4).Formula = "=Sum(" & Range(cells(3, report_col4), cells(clawbacks_pending_row - 1, report_col4)).Address() & ")"
		.cells(clawbacks_pending_row, report_col5).Formula = "=Sum(" & Range(cells(3, report_col5), cells(clawbacks_pending_row - 1, report_col5)).Address() & ")"
		.cells(clawbacks_pending_row, report_col6).Formula = "=Sum(" & Range(cells(3, report_col6), cells(clawbacks_pending_row - 1, report_col6)).Address() & ")"

        .Range("A:F").EntireColumn.AutoFit
		For x = 3 to clawbacks_pending_row
			if .cells(x, report_col4) < 0 then
				.cells(x, report_col4).Font.ColorIndex = 3
			End if
			if .cells(x, report_col5) < 0 then
				.cells(x, report_col5).Font.ColorIndex = 3
			End if
			if .cells(x, report_col6) < 0 then
				.cells(x, report_col6).Font.ColorIndex = 3
			End if
		Next x
    End With
    
    With Worksheets("Jobs in Jeopardy")
		.cells(jobs_in_jeopardy_row, report_systemsize_col) = "Total:"
		.cells(jobs_in_jeopardy_row, report_col5).Formula = "=Sum(" & Range(cells(3, report_col5), cells(jobs_in_jeopardy_row - 1, report_col5)).Address() & ")"
		.cells(jobs_in_jeopardy_row, report_col6).Formula = "=Sum(" & Range(cells(3, report_col6), cells(jobs_in_jeopardy_row - 1, report_col6)).Address() & ")"

        .Range("A:F").EntireColumn.AutoFit
		For x = 3 to jobs_in_jeopardy_row
			if .cells(x, report_col4) < 0 then
				.cells(x, report_col4).Font.ColorIndex = 3
			End if
			if .cells(x, report_col5) < 0 then
				.cells(x, report_col5).Font.ColorIndex = 3
			End if
			if .cells(x, report_col6) < 0 then
				.cells(x, report_col6).Font.ColorIndex = 3
			End if
		Next x
    End With
    
    With Worksheets("Jobs in Progress")
		.cells(jobs_in_progress_row, report_systemsize_col) = "Total:"
		.cells(jobs_in_progress_row, report_col4).Formula = "=Sum(" & Range(cells(3, report_col4), cells(jobs_in_progress_row - 1, report_col4)).Address() & ")"
		.cells(jobs_in_progress_row, report_col5).Formula = "=Sum(" & Range(cells(3, report_col5), cells(jobs_in_progress_row - 1, report_col5)).Address() & ")"
		.cells(jobs_in_progress_row, report_col6).Formula = "=Sum(" & Range(cells(3, report_col6), cells(jobs_in_progress_row - 1, report_col6)).Address() & ")"

        .Range("A:F").EntireColumn.AutoFit
		For x = 3 to jobs_in_progress_row
			if .cells(x, report_col4) < 0 then
				.cells(x, report_col4).Font.ColorIndex = 3
			End if
			if .cells(x, report_col5) < 0 then
				.cells(x, report_col5).Font.ColorIndex = 3
			End if
			if .cells(x, report_col6) < 0 then
				.cells(x, report_col6).Font.ColorIndex = 3
			End if
		Next x
    End With
    
    With Worksheets("Other")
		.cells(other_row, report_col4, report_systemsize_col) = "Total:"
		.cells(other_row, report_col4).Formula = "=Sum(" & Range(cells(3, report_col4), cells(other_row - 1, report_col4)).Address() & ")"
		.cells(other_row, report_col5).Formula = "=Sum(" & Range(cells(3, report_col5), cells(other_row - 1, report_col5)).Address() & ")"

        .Range("A:F").EntireColumn.AutoFit
		For x = 3 to other_row
			if .cells(x, report_col4) < 0 then
				.cells(x, report_col4).Font.ColorIndex = 3
			End if
			if .cells(x, report_col5) < 0 then
				.cells(x, report_col5).Font.ColorIndex = 3
			End if
			if .cells(x, report_col6) < 0 then
				.cells(x, report_col6).Font.ColorIndex = 3
			End if
		Next x
    End With
    
    ThisWorkbook.Sheets(Array("Jobs Installed", "Clawbacks Pending", "Jobs in Jeopardy", "Jobs in Progress", "Other")).Move
    ActiveWorkbook.SaveAs ("C:\users\ezra\desktop\" & "Finance Report" & "\" & rep & ".xlsx")
    ActiveWorkbook.Close
    
    rep_row = rep_row + 1
    
Loop Until Sheets("RepList").Cells(rep_row, 2) = ""

End Sub

Sub format_sheet(ByVal sheetName As String, ByVal Col4 As String, ByVal Col5 As String, ByVal Col6 As String)

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

Sub email(ByVal rep As String)

    'Email portion of the code
        Dim OutApp As Object
        Dim OutMail As Object

        Set OutApp = CreateObject("Outlook.Application")
        Set OutMail = OutApp.CreateItem(0)

        On Error Resume Next
       
        With OutMail
            .To = rep
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

