'Override Report'
'SORT BY Override ID'
Sub report_Overrides

''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Setting Up Sheet''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''

	'instantiate variables'
	Dim iCurrentID, iPreviousID AS Integer
	Dim currentRow As Integer

	Dim masterReport, overrideMaster As Workbook
	Dim payments, master as Worksheet

	iPreviousID = 0
	iCurrentID = 1

	FilePath = Application.GetOpenFilename()
    FileName = convertToName(FilePath)

    If isWorkBookOpen(FilePath) Then
        Set masterReport = Workbooks(FileName)
    Else
        Set masterReport = Workbooks.Open(FilePath)
    End If

    'set sheets'
    Set overrideMaster = Workbooks(ThisWorkbook.Name)
    Set master         = masterReport.Sheets("Current Data")
    Set payments       = Master.Sheets("Payments")

    Set histSheet      = Master.Sheets("Override Past")



''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''Begin Calculations''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''

    'Loop through sorted payments tab to sum override IDs'
    Do Until IsEmpty(master.Cells(currentRow, 1).Value)
    	
    	'if we are no longer dealing with the same OverrideID'
    	If iPreviousID <> iCurrentID Then
    		'print out previous override ID to report/Reset Variables'

    	Else
    		'sum current row with last row'

    	End If

    Loop

End Sub

'Prints out overrideID to report'
Sub print(ByRef currentRow As Integer)

End Sub