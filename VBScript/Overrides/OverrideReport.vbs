'Override Report'

	'Report Columns'
	Dim repCol, customerCol, jobCol, kWCol, typeCol, _
	 rateCol, earnedCol, paidCol, dueCol As Integer

	 'Override Master Columns' 
	Dim overIDCol_pay, overRepCol_pay, repCol_pay, customerCol_pay, jobCol_pay, _
	 typeCol_pay, rateCol_pay,  kWCol_pay, paidCol_pay As Integer

	 'Evolve Report Columns'
	Dim CustomerCol_ev,  JobIDCol_ev, kWCol_ev, StatusCol_ev,
     SubStatusCol_ev, theDateCol_ev, repEmailCol_ev As Integer



'SORT BY Override ID'
Sub report_Overrides

''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Setting Up Sheet''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''

	'declare variables'
	Dim iCurrentID, iPreviousID AS Integer
	Dim currentRow As Integer
	Dim reportName As String

	Dim EvolveReport, OverMaster As Workbook
	Dim payments, curentData, report as Worksheet

	'initialize column variables'
	call initVar

	iPreviousID = 0
	iCurrentID = 1

	FilePath = Application.GetOpenFilename()
    FileName = convertToName(FilePath)

    If isWorkBookOpen(FilePath) Then
        Set EvolveReport = Workbooks(FileName)
    Else
        Set EvolveReport = Workbooks.Open(FilePath)
    End If

    'name of report'
    reportName = Date() & " Report"

    'set sheets'
    Set OverMaster     = Workbooks(ThisWorkbook.Name)
    Set currentData    = EvolveReport.Sheets("Current Data")
    Set payments       = OverMaster.Sheets("Payments")



''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Begin Calculations''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''

	'create report'
	call createReport reportName

    'Loop through sorted payments tab to sum override IDs'
    Do Until IsEmpty(payments.Cells(currentRow, 1).Value)
    	
    	'if we are no longer dealing with the same OverrideID'
    	If iPreviousID <> iCurrentID Then
    		'print out previous override ID to report/Reset Variables'
    		call print currentRow
    	Else
    		'sum current row with last row'

    	End If

    Loop

End Sub

'Prints out overrideID to report'
Sub print(ByRef currentRow As Integer)

End Sub

'create and format report sheet'
Sub createReport(ByRef sheetName As String)
	

	Worksheets.Add(, Worksheets(Worksheets.Count)).Name = sheetName

    With Worksheets(sheetName)
        .Cells(1, 1)           = sheetName
        .Cells(2, repCol)      = "Rep"
        .Cells(2, customerCol) = "Customer"
        .Cells(2, jobCol)      = "JobID"
        .Cells(2, kWCol)       = "System Size"
        .Cells(2, typeCol)     = "OverrideType"
        .Cells(2, rateCol)     = "Rate"
        .Cells(2, earnedCol)   = "Earned"
        .Cells(2, paidCol)     = "Paid"
        .Cells(2, dueCol)      = "Due"

    End With

    'Formats the main header
    With Worksheets(sheetName).Range("A1:I1")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 77, 0)
        .Font.Color = RGB(255, 255, 255)
    End With

    'Formats the column headers
    With Worksheets(sheetName).Range("A2:I2")
        .HorizontalAlignment = xlLeft
        .Font.Bold = True
        .Interior.Color = RGB(0, 0, 0)
        .Font.Color = RGB(255, 255, 255)
    End With
End Sub

Sub initVar()
	'report Columns'
	repCol      = 1
	customerCol = 2
	jobCol      = 3
	kWCol       = 4
	typeCol     = 5
	rateCol     = 6
	earnedCol   = 7
	paidCol     = 8
	duecol      = 9

	'payments columns'
	overIDCol_pay   = 1
	overRepCol_pay  = 2
	repCol_pay      = 4
	customerCol_pay = 6
	jobCol_pay      = 7
	typeCol_pay     = 13
	rateCol_pay     = 14
	kWCol_pay       = 15
	paidCol_pay     = 16

	'evolve master report columns'
	CustomerCol_ev  = 1
	JobIDCol_ev     = 2
	kWCol_ev        = 3
	StatusCol_ev    = 4
	SubStatusCol_ev = 5
	theDateCol_ev   = 7
	repEmailCol_ev  = 17

End Sub