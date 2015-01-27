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
dim rate_col as integer

rate_col = 5
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
	      
    rep = Sheets("RepList").Cells(rep_row, 2)
    rep_name = Sheets("RepList").Cells(rep_row, 1)
	rep_ID = sheets("RepList").cells(rep_row, 3)
	rate = Sheets("RepList").cells(rep_row, 4)
    
    'cycles through the accounts
    Do
        If Sheets("Current Data").Cells(current_data_row, 17) = rep Then
        
        'assigns values to the variables reported on
        job_id = Sheets("Current Data").Cells(current_data_row, 2)
        customer = Sheets("Current Data").Cells(current_data_row, 1)
        system_size = Sheets("Current Data").Cells(current_data_row, 3)
        
        job_in_jeopardy = Sheets("Current Data").Cells(current_data_row, 19)
        cancellation = Sheets("Current Data").Cells(current_data_row, 20)
        installed = Sheets("Current Data").Cells(current_data_row, 21)
        in_progress = Sheets("Current Data").Cells(current_data_row, 22)
        
        status = Sheets("Current Data").Cells(current_data_row, 4)

        'code for calculating earned based on rep. This is in here for confidentiality.
		with sheets("Financial Data")
			do until (IsEmpty(.cells(financial_data_row, 2)))
				if .cells(financial_data_row, 2) = rep_ID AND .cells(financial_data_row, 3) = customer then
					earned = system_size * sheets("RepList").cells(rep_row, rate_col)
					financial_data_row = financial_data_row + 1
				Else
					financial_data_row = financial_data_row + 1
				End if
			Loop
		end with
        
        'code for summing the amount paid for the account
        Dim sale_amount_row As Integer
        sale_amount_row = 2
        
        Do
            If Sheets("Financial Data").Cells(sale_amount_row, 2) = rep_name AND sheets("Financial Data").cells(sale_amount_row, 7) = job_id Then
                sale_amount = sale_amount + Sheets("Financial Data").Cells(sale_amount_row, 5)
                sale_amount_row = sale_amount_row + 1
            Else
                sale_amount_row = sale_amount_row + 1
            End If
        Loop Until Sheets("Financial Data").Cells(sale_amount_row, 3) = ""

        'Determines which reporting sheet the account goes to
        If installed = "TRUE" Or installed = "True" Then
            Sheets("Commissions Earned").Cells(commissions_earned_row, 1) = customer
            Sheets("Commissions Earned").Cells(commissions_earned_row, 2) = job_id
            Sheets("Commissions Earned").Cells(commissions_earned_row, 3) = system_size
            Sheets("Commissions Earned").Cells(commissions_earned_row, 4) = earned
            Sheets("Commissions Earned").Cells(commissions_earned_row, 5) = sale_amount
            Sheets("Commissions Earned").Cells(commissions_earned_row, 6) = earned - sale_amount
            
            commissions_earned_row = commissions_earned_row + 1
        ElseIf cancellation = "TRUE" Or cancellation = "True" Then
            Sheets("Clawbacks Pending").Cells(clawbacks_pending_row, 1) = customer
            Sheets("Clawbacks Pending").Cells(clawbacks_pending_row, 2) = job_id
            Sheets("Clawbacks Pending").Cells(clawbacks_pending_row, 3) = system_size
            Sheets("Clawbacks Pending").Cells(clawbacks_pending_row, 4) = earned
            Sheets("Clawbacks Pending").Cells(clawbacks_pending_row, 5) = sale_amount
            Sheets("Clawbacks Pending").Cells(clawbacks_pending_row, 6) = earned - sale_amount
            
            clawbacks_pending_row = clawbacks_pending_row + 1
        ElseIf job_in_jeopardy = "TRUE" Or job_in_jeopardy = "True" Then
            Sheets("Jobs in Jeopardy").Cells(jobs_in_jeopardy_row, 1) = customer
            Sheets("Jobs in Jeopardy").Cells(jobs_in_jeopardy_row, 2) = job_id
            Sheets("Jobs in Jeopardy").Cells(jobs_in_jeopardy_row, 3) = system_size
            Sheets("Jobs in Jeopardy").Cells(jobs_in_jeopardy_row, 4) = status
            Sheets("Jobs in Jeopardy").Cells(jobs_in_jeopardy_row, 5) = sale_amount
            Sheets("Jobs in Jeopardy").Cells(jobs_in_jeopardy_row, 6) = 0 - sale_amount
            
            jobs_in_jeopardy_row = jobs_in_jeopardy_row + 1
        ElseIf in_progress = "TRUE" Or in_progress = "True" Then
            Sheets("Jobs in Progress").Cells(jobs_in_progress_row, 1) = customer
            Sheets("Jobs in Progress").Cells(jobs_in_progress_row, 2) = job_id
            Sheets("Jobs in Progress").Cells(jobs_in_progress_row, 3) = system_size
            Sheets("Jobs in Progress").Cells(jobs_in_progress_row, 4) = earned
            Sheets("Jobs in Progress").Cells(jobs_in_progress_row, 5) = sale_amount
            Sheets("Jobs in Progress").Cells(jobs_in_progress_row, 6) = earned - sale_amount
            
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

