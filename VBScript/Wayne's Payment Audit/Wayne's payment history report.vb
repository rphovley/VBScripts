Sub WaynesPayDateAudit()

'this is where I delineate variables
Dim weekly_audit As Workbook
Dim Wayne As Worksheet
Dim Historical As Worksheet
Dim Lists As Worksheet
Dim Report As Worksheet

Set weekly_audit = Workbooks("Waynes historical pay.xlsm")
Set Wayne = weekly_audit.Sheets("Wayne")
Set Historical = weekly_audit.Sheets("Historical")
Set Lists = weekly_audit.Sheets("Lists")
Set Report = weekly_audit.Sheets("Report")

Dim wayne_row As Integer
    wayne_row = 2
Dim historical_row As Integer
    historical_row = 2
Dim rep_id_row As Integer
    rep_id_row = 2
Dim date_row As Integer
    date_row = 2
Dim report_row As Integer
    report_row = 2

Dim deposit_date_col As Integer
    deposit_date_col = 1
Dim payment_amount_col As Integer
    payment_amount_col = 5

Dim list_date_col As Integer
    list_date_col = 3
Dim list_rep_col As Integer
    list_rep_col = 1
    
Dim rep_id As Variant  'I chose variant because sometimes the data has ??? or #N/A which throws out an error if it is an integer type
Dim deposit_date As Date
Dim end_date As Date
Dim payment_sum As Currency

Dim final_row As Long
    final_row = Historical.UsedRange.Rows.Count
    
Do Until IsEmpty(Lists.Cells(date_row, list_date_col))
    payment_sum = 0
    end_date = Lists.Cells(date_row, list_date_col)
    deposit_date = end_date
    rep_id_row = 2
    
    Do Until IsEmpty(Lists.Cells(rep_id_row, list_rep_col))
        rep_id = Lists.Cells(rep_id_row, list_rep_col)
        payment_sum = 0
        historical_row = 2
        For x = 2 To final_row
            If Historical.Cells(historical_row, 2) = rep_id Then
                If Historical.Cells(historical_row, deposit_date_col) = deposit_date Then
                    payment_sum = payment_sum + Historical.Cells(historical_row, payment_amount_col)
                End If
            End If
            
            historical_row = historical_row + 1
        Next x
        'Before it goes to the next rep, that rep's info needs to be printed to the Report sheet
        If payment_sum <> 0 Then
            With Report
                .Cells(report_row, 1) = rep_id
                .Cells(report_row, 2) = end_date
                .Cells(report_row, 3) = payment_sum
            End With
            report_row = report_row + 1
        End If
        
        rep_id_row = rep_id_row + 1
    Loop
    
    date_row = date_row + 1
Loop
    
End Sub
