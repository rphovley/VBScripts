'Override Report'

	'Report Columns'
	Dim repCol, customerCol, jobCol, kWCol, overTypeCol, _
	 rateCol, earnedCol, paidCol, dueCol As Integer

	 'Override Master Columns' 
	Dim overIDCol_pay, overRepCol_pay, repCol_pay, customerCol_pay, jobCol_pay, _
	 overTypeCol_pay, rateCol_pay,  kWCol_pay, paidCol_pay As Integer

	 'Evolve Report Columns'
	Dim CustomerCol_ev,  JobIDCol_ev, kWCol_ev, StatusCol_ev, _
     SubStatusCol_ev, theDateCol_ev, repEmailCol_ev As Integer

    'Report Values'
    Dim repName, customer, job, overType As String
    Dim kW, rate, earned, paid, due As Double



'SORT BY Override ID'
Sub report_Overrides

''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Setting Up Sheet''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''

	'declare variables'
	Dim iCurrentID, iPreviousID AS Integer
	Dim currentRow_pay, currentRow_report As Integer
	Dim reportName, overMasterFileName, EvolveReportFileName As String

	Dim EvolveReport, OverMaster As Workbook
	Dim payments, curentData, report as Worksheet

	'initialize column variables'
	call initVar

	iPreviousID = 0
	iCurrentID = 1
	currentRow_pay = 2
	currentRow_report = 3

	FilePath = Application.GetOpenFilename()
    FileName = convertToName(FilePath)
    EvolveReportFileName = FileName
    overMasterFileName = ThisWorkbook.Name

    Set EvolveReport = Workbooks(FileName)

    'name of report'
    reportName = Format(Date(), "dd-mm-yyyy") & " Report"

    'set sheets'
    Set OverMaster     = Workbooks(ThisWorkbook.Name)
    Set currentData    = EvolveReport.Sheets("Current Data")
    Set payments       = OverMaster.Sheets("Payments")



''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Begin Calculations''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''

	'create report'
	createReport reportName

    'Loop through sorted payments tab to sum override IDs'
    Do Until IsEmpty(payments.Cells(currentRow_pay, 1).Value)
    	
    	'if we are no longer dealing with the same OverrideID'
    	If iPreviousID <> iCurrentID Then
    		'print out previous override ID to report/Reset Variables'
    		printReport reportName, currentRow_pay, currentRow_report
    		currentRow_report = currentRow_report + 1
    	Else
    		'sum current row with last row'
    		paid = paid + payments.Cells(currentRow_pay, paidCol_pay).Value

    	End If

    	iPreviousID = iCurrentID
    	currentRow_pay = currentRow_pay + 1
    	iCurrentID = payments.Cells(currentRow_pay, 1).Value
    Loop

End Sub

'Prints out overrideID to report'
Sub printReport(ByVal sheetName As String, ByVal currentRow As Integer, ByVal currentRow_report As Integer)
	With Worksheets(sheetName)
        .Cells(currentRow_report, repCol)      = repName
        .Cells(currentRow_report, customerCol) = customer
        .Cells(currentRow_report, jobCol)      = job
        .Cells(currentRow_report, kWCol)       = kW
        .Cells(currentRow_report, overTypeCol) = overType
        .Cells(currentRow_report, rateCol)     = rate
        .Cells(currentRow_report, earnedCol)   = earned
        .Cells(currentRow_report, paidCol)     = paid
        .Cells(currentRow_report, dueCol)      = due
	End With

	call ResetVar

End Sub

'create and format report sheet'
Sub createReport(ByVal sheetName As String)
	

	Worksheets.Add(, Worksheets(Worksheets.Count)).Name = sheetName
	call ResetVar
    With Worksheets(sheetName)
        .Cells(1, 1)           = sheetName
        .Cells(2, repCol)      = "Rep"
        .Cells(2, customerCol) = "Customer"
        .Cells(2, jobCol)      = "JobID"
        .Cells(2, kWCol)       = "System Size"
        .Cells(2, overTypeCol)     = "OverrideoverType"
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

'sub to initialize column variables'
Sub initVar()
	'report Columns'
	repCol      = 1
	customerCol = 2
	jobCol      = 3
	kWCol       = 4
	overTypeCol     = 5
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
	overTypeCol_pay     = 13
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

'Reset Report Variables'
Sub ResetVar()
	repName  = ""
    customer = ""
    job      = ""
    kW       = 0
    overType     = ""
    rate     = 0
    earned   = 0
    paid     = 0
    due      = 0
End Sub

Function convertToName(ByVal Path As String) As String

     For Each wbk1 In Workbooks
        If (wbk1.Path & "\" & wbk1.Name = Path) Then
            convertToName = wbk1.Name
            Exit For
        End If
    Next


End Function
