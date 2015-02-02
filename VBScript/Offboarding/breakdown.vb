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
dim payout_percent as double


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
	dim payout_percent_col as integer
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
    
    Call format_sheet("Other", "System Value", "Amount", "Pending")
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    current_data_row = 2
    commissions_earned_row = 3
    clawbacks_pending_row = 3
    jobs_in_jeopardy_row = 3
    jobs_in_progress_row = 3
    other_row = 3
	payout_percent_col = 26

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
				payout_percent = .cells(current_data_row, payout_percent_col)
                
                job_in_jeopardy = .Cells(current_data_row, current_jeopardy_col)
                cancellation = .Cells(current_data_row, current_cancellation_col)
                installed = .Cells(current_data_row, current_installed_col)
                in_progress = .Cells(current_data_row, current_inprogress_col)
            End With
            
            status = Sheets("Current Data").Cells(current_data_row, current_status_col)

            'code for calculating earned based on rep.
			If payout_percent <> "" or payout_percent = 0 then
				system_value = system_size * rate * payout_percent
			Else
				system_value = system_size * rate * 1
			End if

        
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
                    If .Cells(financial_row, 1) = "Pending" Then
                        customer = .Cells(financial_row, financial_customer_col)
                        job_id = 0
						'.Cells(financial_row, financial_jobID_col)
                        system_size = .Cells(financial_row, financial_kW_col)
                        system_value = system_size * rate
                        sale_amount = .Cells(financial_row, financial_payment_col)
                        
                        With Sheets("Other")
                            .Cells(other_row, report_customer_col) = customer
                            .Cells(other_row, report_jobID_col) = job_id
                            .Cells(other_row, report_systemsize_col) = system_size
                            .Cells(other_row, report_col4) = system_value
                            .Cells(other_row, report_col6) = sale_amount
                        End With
                    
                    Else
                        customer = .Cells(financial_row, financial_customer_col)
                        job_id = .Cells(financial_row, financial_jobID_col)
                        system_size = 0
						'.Cells(financial_row, financial_kW_col)
                        system_value = system_size * rate
                        sale_amount = .Cells(financial_row, financial_payment_col)
                        
                        With Sheets("Other")
                            .Cells(other_row, report_customer_col) = customer
                            .Cells(other_row, report_jobID_col) = job_id
                            .Cells(other_row, report_systemsize_col) = system_size
                            .Cells(other_row, report_col4) = system_value
                            .Cells(other_row, report_col5) = sale_amount
                        End With
                    End If
                    other_row = other_row + 1
                End If
                
                financial_row = financial_row + 1
            Else
                financial_row = financial_row + 1
            End If
        Loop
    End With
    
    With Worksheets("Jobs Installed")
        .Cells(commissions_earned_row, report_systemsize_col) = "Total:"
        .Cells(commissions_earned_row, report_col4).Formula = "=Sum(" & Range(Cells(3, report_col4), Cells(commissions_earned_row - 1, report_col4)).Address() & ")"
        .Cells(commissions_earned_row, report_col5).Formula = "=Sum(" & Range(Cells(3, report_col5), Cells(commissions_earned_row - 1, report_col5)).Address() & ")"
        .Cells(commissions_earned_row, report_col6).Formula = "=Sum(" & Range(Cells(3, report_col6), Cells(commissions_earned_row - 1, report_col6)).Address() & ")"

        .Range("A:F").EntireColumn.AutoFit
        For x = 3 To commissions_earned_row
            If .Cells(x, report_col4) < 0 Then
                .Cells(x, report_col4).Font.ColorIndex = 3
            End If
            If .Cells(x, report_col5) < 0 Then
                .Cells(x, report_col5).Font.ColorIndex = 3
            End If
            If .Cells(x, report_col6) < 0 Then
                .Cells(x, report_col6).Font.ColorIndex = 3
            End If
        Next x
    End With
    
    With Worksheets("Clawbacks Pending")
        .Cells(clawbacks_pending_row, report_systemsize_col) = "Total:"
        .Cells(clawbacks_pending_row, report_col4).Formula = "=Sum(" & Range(Cells(3, report_col4), Cells(clawbacks_pending_row - 1, report_col4)).Address() & ")"
        .Cells(clawbacks_pending_row, report_col5).Formula = "=Sum(" & Range(Cells(3, report_col5), Cells(clawbacks_pending_row - 1, report_col5)).Address() & ")"
        .Cells(clawbacks_pending_row, report_col6).Formula = "=Sum(" & Range(Cells(3, report_col6), Cells(clawbacks_pending_row - 1, report_col6)).Address() & ")"

        .Range("A:F").EntireColumn.AutoFit
        For x = 3 To clawbacks_pending_row
            If .Cells(x, report_col4) < 0 Then
                .Cells(x, report_col4).Font.ColorIndex = 3
            End If
            If .Cells(x, report_col5) < 0 Then
                .Cells(x, report_col5).Font.ColorIndex = 3
            End If
            If .Cells(x, report_col6) < 0 Then
                .Cells(x, report_col6).Font.ColorIndex = 3
            End If
        Next x
    End With
    
    With Worksheets("Jobs in Jeopardy")
        .Cells(jobs_in_jeopardy_row, report_systemsize_col) = "Total:"
        .Cells(jobs_in_jeopardy_row, report_col5).Formula = "=Sum(" & Range(Cells(3, report_col5), Cells(jobs_in_jeopardy_row - 1, report_col5)).Address() & ")"
        .Cells(jobs_in_jeopardy_row, report_col6).Formula = "=Sum(" & Range(Cells(3, report_col6), Cells(jobs_in_jeopardy_row - 1, report_col6)).Address() & ")"

        .Range("A:F").EntireColumn.AutoFit
        For x = 3 To jobs_in_jeopardy_row
            If .Cells(x, report_col4) < 0 Then
                .Cells(x, report_col4).Font.ColorIndex = 3
            End If
            If .Cells(x, report_col5) < 0 Then
                .Cells(x, report_col5).Font.ColorIndex = 3
            End If
            If .Cells(x, report_col6) < 0 Then
                .Cells(x, report_col6).Font.ColorIndex = 3
            End If
        Next x
    End With
    
    With Worksheets("Jobs in Progress")
        .Cells(jobs_in_progress_row, report_systemsize_col) = "Total:"
        .Cells(jobs_in_progress_row, report_col4).Formula = "=Sum(" & Range(Cells(3, report_col4), Cells(jobs_in_progress_row - 1, report_col4)).Address() & ")"
        .Cells(jobs_in_progress_row, report_col5).Formula = "=Sum(" & Range(Cells(3, report_col5), Cells(jobs_in_progress_row - 1, report_col5)).Address() & ")"
        .Cells(jobs_in_progress_row, report_col6).Formula = "=Sum(" & Range(Cells(3, report_col6), Cells(jobs_in_progress_row - 1, report_col6)).Address() & ")"

        .Range("A:F").EntireColumn.AutoFit
        For x = 3 To jobs_in_progress_row
            If .Cells(x, report_col4) < 0 Then
                .Cells(x, report_col4).Font.ColorIndex = 3
            End If
            If .Cells(x, report_col5) < 0 Then
                .Cells(x, report_col5).Font.ColorIndex = 3
            End If
            If .Cells(x, report_col6) < 0 Then
                .Cells(x, report_col6).Font.ColorIndex = 3
            End If
        Next x
    End With
    
    With Worksheets("Other")
        .Cells(other_row, report_systemsize_col) = "Total:"
        .Cells(other_row, report_col4).Formula = "=Sum(" & Range(Cells(3, report_col4), Cells(other_row - 1, report_col4)).Address() & ")"
        .Cells(other_row, report_col5).Formula = "=Sum(" & Range(Cells(3, report_col5), Cells(other_row - 1, report_col5)).Address() & ")"
        .Cells(other_row, report_col6).Formula = "=Sum(" & Range(Cells(3, report_col6), Cells(other_row - 1, report_col6)).Address() & ")"

        .Range("A:F").EntireColumn.AutoFit
        For x = 3 To other_row
            If .Cells(x, report_col4) < 0 Then
                .Cells(x, report_col4).Font.ColorIndex = 3
            End If
            If .Cells(x, report_col5) < 0 Then
                .Cells(x, report_col5).Font.ColorIndex = 3
            End If
            If .Cells(x, report_col6) < 0 Then
                .Cells(x, report_col6).Font.ColorIndex = 3
            End If
        Next x
    End With
    
    ThisWorkbook.Sheets(Array("Jobs Installed", "Clawbacks Pending", "Jobs in Jeopardy", "Jobs in Progress", "Other")).Move
    ActiveWorkbook.SaveAs ("C:\users\ezra\desktop\" & "Finance Report" & "\" & rep & ".xlsx")
    ActiveWorkbook.Close
    
    Call email(rep)
    
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
            .Body = "Hello," & vbNewLine & vbNewLine _
            & "The following attachment is a list of all your jobs with Evolve Solar with how much you received per job and when." & vbNewLine & vbNewLine _
            & "Also, the following is a description of how to read the attachment." & vbNewLine & vbNewLine _
            & "'Jobs Installed'" & vbNewLine & vbNewLine _
            & Chr(149) & "This shows all SolarCity jobs that have hit installation and how much you've been paid." & vbNewLine _
            & Chr(149) & "The 'Due': If Black- That means we owe the following amount. If Red- You were overpaid (most likely from downsizing) and will be taken from future installs." & vbNewLine & vbNewLine _
            & "'Clawbacks Pending'" & vbNewLine & vbNewLine _
            & Chr(149) & "This shows all CANCELLED jobs." & vbNewLine _
            & Chr(149) & "If there is a number (black) in the column 'Due to Evolve,' that money is still owed to Evolve and will be taken from future installs." & vbNewLine & vbNewLine _
            & "'Jobs in Jeopardy'" & vbNewLine & vbNewLine _
            & Chr(149) & "This shows all jobs that are in jeopardy and shows the amount that IF cancelled, the money is owed back to Evolve." & vbNewLine & vbNewLine _
            & "'Jobs in Progress'" & vbNewLine & vbNewLine _
            & Chr(149) & "This shows all jobs that are still in progress but have yet to hit installation." & vbNewLine _
            & Chr(149) & "The 'Paid' column shows how much you've been paid per job." & vbNewLine _
            & Chr(149) & "The 'Potentially Due' column shows how much money can still be made once the account hits installation." & vbNewLine & vbNewLine _
            & "'Other'" & vbNewLine & vbNewLine _
            & Chr(149) & "This shows all other payments that have occurred, example, Overrides, special advances, reimbursements, rent, etc." & vbNewLine _
            & Chr(149) & "The 'Pending' column means that these line items have not yet hit your paycheck. Reasons being is either you have a negative balance with Evolve, etc." & vbNewLine _
            & Chr(149) & "If the 'Pending' column is positive, Evolve owes you the money and if the account is negative, you still owe Evolve." & vbNewLine & vbNewLine _
            & "We are happy to talk with you about this (Please do not respond to this email. Use the financial inquiry below). As you know, we are in development to make this information live online and updated weekly. Since the web version is not yet available, but close, we are sending personal emails to each of you as we promised we would get this information to you by this week." & vbNewLine & vbNewLine _
            & "Thank you," & vbNewLine & vbNewLine _
            & "The Finance Department" & vbNewLine & vbNewLine _
            & "http://goo.gl/forms/j0fQXnf5ov"
            
            .Attachments.Add ("C:\users\ezra\desktop\" & "Finance Report" & "\" & rep & ".xlsx")
            '.Display
            .Send
        End With
        On Error GoTo 0

        Set OutMail = Nothing
        Set OutApp = Nothing
    
End Sub


