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

rep_row = 2
Do
'Creates a new tab


Worksheets.Add(, Worksheets(Worksheets.Count)).Name = "Jobs in Jeopardy"
Worksheets.Add(, Worksheets(Worksheets.Count)).Name = "Jobs in Progress"


'Formats the new spreadsheets
    'Commissions Earned sheet
    

    'Clawbacks Pending sheet
    Worksheets("Clawbacks Pending").Cells(1, 1) = "Clawbacks Pending"
    Worksheets("Clawbacks Pending").Cells(2, 1) = "Customer"
    Worksheets("Clawbacks Pending").Cells(2, 2) = "JobID"
    Worksheets("Clawbacks Pending").Cells(2, 3) = "System Size"
    Worksheets("Clawbacks Pending").Cells(2, 4) = "Earned"
    Worksheets("Clawbacks Pending").Cells(2, 5) = "Paid but Never Clawed Back"
    Worksheets("Clawbacks Pending").Cells(2, 6) = "Due to Evolve"

    'Formats the main header
    With Worksheets("Clawbacks Pending").Range("A1:F1")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 77, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    'Formats the column headers
    With Worksheets("Clawbacks Pending").Range("A2:F2")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    'Jobs in Jeopardy sheet
    Worksheets("Jobs in Jeopardy").Cells(1, 1) = "Jobs in Jeopardy"
    Worksheets("Jobs in Jeopardy").Cells(2, 1) = "Customer"
    Worksheets("Jobs in Jeopardy").Cells(2, 2) = "JobID"
    Worksheets("Jobs in Jeopardy").Cells(2, 3) = "System Size"
    Worksheets("Jobs in Jeopardy").Cells(2, 4) = "Status"
    Worksheets("Jobs in Jeopardy").Cells(2, 5) = "Paid"
    Worksheets("Jobs in Jeopardy").Cells(2, 6) = "High Probability of Cancelling"

    'Formats the main header
    With Worksheets("Jobs in Jeopardy").Range("A1:F1")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 77, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    'Formats the column headers
    With Worksheets("Jobs in Jeopardy").Range("A2:F2")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With
    
    'Jobs in Progress sheet
    Worksheets("Jobs in Progress").Cells(1, 1) = "Jobs in Progress"
    Worksheets("Jobs in Progress").Cells(2, 1) = "Customer"
    Worksheets("Jobs in Progress").Cells(2, 2) = "JobID"
    Worksheets("Jobs in Progress").Cells(2, 3) = "System Size"
    Worksheets("Jobs in Progress").Cells(2, 4) = "Potentially Earned"
    Worksheets("Jobs in Progress").Cells(2, 5) = "Paid"
    Worksheets("Jobs in Progress").Cells(2, 6) = "Potentially Due"

    'Formats the main header
    With Worksheets("Jobs in Progress").Range("A1:F1")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 77, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    'Formats the column headers
    With Worksheets("Jobs in Progress").Range("A2:F2")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With
''''''''''''''''''''''''''''''''''''''''''''''''''''''
    financial_data_row = 2
    current_data_row = 2
    commissions_earned_row = 3
    clawbacks_pending_row = 3
    jobs_in_jeopardy_row = 3
    jobs_in_progress_row = 3
      
    rep = Sheets("RepList").Cells(rep_row, 2)
    rep_name = Sheets("RepList").Cells(rep_row, 1)
    
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
        If rep = "alejandro@evolvesolar.com" Then
            earned = system_size * 250
        ElseIf rep = "benjamin@evolvesolar.com" Then
            earned = system_size * 200
        ElseIf rep = "daniel.field@evolvesolar.com" Then
            earned = system_size * 250
        ElseIf rep = "heriberto@evolvesolar.com" Then
            earned = system_size * 200
        ElseIf rep = "stoxen@evolvesolar.com" Then
            earned = system_size * 200
        ElseIf rep = "kent.shumway@evolvesolar.com" Then
            earned = system_size * 200
        ElseIf rep = "jaredb@evolvesolar.com" Then
            earned = system_size * 250
        ElseIf rep = "scottk@evolvesolar.com" Then
            earned = system_size * 250
        ElseIf rep = "ryanp@evolvesolar.com" Then
            earned = system_size * 250
        ElseIf rep = "tyson@evolvesolar.com" Then
            earned = system_size * 200
        ElseIf rep = "jasonm@evolvesolar.com" Then
            earned = system_size * 200
        ElseIf rep = "matt@evolvesolar.com" Then
            earned = system_size * 280
        End If
        
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

sub commissions_earned()

Worksheets.Add(, Worksheets(Worksheets.Count)).Name = "Commissions Earned"

Worksheets("Commissions Earned").Cells(1, 1) = "Commissions Earned"
    Worksheets("Commissions Earned").Cells(2, 1) = "Customer"
    Worksheets("Commissions Earned").Cells(2, 2) = "JobID"
    Worksheets("Commissions Earned").Cells(2, 3) = "System Size"
    Worksheets("Commissions Earned").Cells(2, 4) = "Earned"
    Worksheets("Commissions Earned").Cells(2, 5) = "Paid"
    Worksheets("Commissions Earned").Cells(2, 6) = "Due"

    'Formats the main header
    With Worksheets("Commissions Earned").Range("A1:F1")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 77, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    'Formats the column headers
    With Worksheets("Commissions Earned").Range("A2:F2")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

End Sub

Sub clawbacks_pending()
Worksheets.Add(, Worksheets(Worksheets.Count)).Name = "Clawbacks Pending"


End Sub
