	'Columns for the "Report Tab"'
	Dim repJobIdCol, repDateCol, repkWCol, repStatusCol, _
	 repOldNewCol, repPaidOutCol, repCurValCol, repEstCol, _
	 repActCol, repCheckCol As Integer

	'Columns for the "Master Report" Tab'
	Dim masJobIdCol, masDateCol, maskWCol, masStatusCol, _
	 masFinalCol, masInstallCol, masCancelledCol As Integer


	 'Collection KEYS'
	Dim dJOBID, dKW, dSTATUS, dDATE, dFINAL, dINSTALL, dCANCELLED AS String
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''Main Sub for Estimate'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Main sub for creating estimate'
'Works with the "SolarCity Audit.xlsm" file in the Historical Breakdowns folder'
Sub createEstimate()

	'initialize variables from above'
	initVar

	'dim the row vars for both tabs'
	Dim MasterReportRow, ReportRow As Integer

	MasterReportRow = 2
	ReportRow       = 2

	'used to pass information back and forth from functions'
	Dim dataFromMasterReport As New Collection

	'Loop through the Master Report'
	Do Until isEmpty(Sheets("Master Report").Cells(MasterReportRow, 1).Value)
	
		'Collect Data from Master Report and Determine what should be paid out to us in the Master Report'
		Set dataFromMasterReport = determinePayout(dataFromMasterReport, MasterReportRow)
		
		'print out what should be paid out in the Report Tab'
	 	printData dataFromMasterReport, ReportRow
		
		Call check_structure(ReportRow, repDateCol, repOldNewCol)
		
	 	'In order to reset the values in a collection the values have to be removed first, this function does that'
		Set dataFromMasterReport = refreshCollection(dataFromMasterReport)

	 	MasterReportRow = MasterReportRow + 1
	 	ReportRow       = ReportRow + 1
	Loop


End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''Supporting Subs and Functions''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Determine what should be paid out to us in the Master Report'
Function determinePayout(ByRef dataFromMasterReport As Collection, ByVal MasterReportRow As Integer) As Collection

	'Collection data from Master Report'
	Set	dataFromMasterReport = setCollection(dataFromMasterReport, MasterReportRow)

	'We make some decision based on what we find in the report'

	'This is returning the collection to the calling sub'
	
	Set determinePayout = dataFromMasterReport
End Function

'Set Collection Values for the data from the Master Report'
Function setCollection(ByRef dataFromMasterReport As Collection, ByVal MasterReportRow As Integer) As Collection

	With Sheets("Master Report")
		dataFromMasterReport.Add .Cells(MasterReportRow, masJobIdCol), dJOBID
	    dataFromMasterReport.Add .Cells(MasterReportRow, maskWCol), dKW
	    dataFromMasterReport.Add .Cells(MasterReportRow, masStatusCol), dSTATUS
	    dataFromMasterReport.Add .Cells(MasterReportRow, masDateCol), dDATE
	    dataFromMasterReport.Add .Cells(MasterReportRow, masFinalCol), dFINAL
	    dataFromMasterReport.Add .Cells(MasterReportRow, masCancelledCol), dCANCELLED
	    dataFromMasterReport.Add .Cells(MasterReportRow, masInstallCol), dINSTALL
	End With

	'this is returning the collection from the calling function'
	Set setCollection = dataFromMasterReport

End Function

'In order to reset the values in a collection the values have to be removed first, this function does that'
Function refreshCollection(ByRef dataFromMasterReport As Collection) As Collection
	
	dataFromMasterReport.Remove dJOBID
    dataFromMasterReport.Remove dKW
    dataFromMasterReport.Remove dSTATUS
    dataFromMasterReport.Remove dDATE
    dataFromMasterReport.Remove dFINAL
    dataFromMasterReport.Remove dCANCELLED
    dataFromMasterReport.Remove dINSTALL

	Set refreshCollection = dataFromMasterReport
End Function
'Sub to print out data gathered into the Report Tab'
Sub printData(ByRef dataFromMasterReport, ByVal ReportRow As Integer)
	
	With Sheets("Report")
		.Cells(ReportRow, repJobIdCol).Value  = dataFromMasterReport.Item(dJOBID)
		.Cells(ReportRow, repDateCol).Value   = dataFromMasterReport.Item(dDATE)
		.Cells(ReportRow, repkWCol).Value     = dataFromMasterReport.Item(dKW)
		.Cells(ReportRow, repStatusCol).Value = dataFromMasterReport.Item(dSTATUS)
		' .Cells(ReportRow, repOldNewCol).Value = dataFromMasterReport.Item(dOLDNEw)
		' .Cells(ReportRow, repEstCol).Value    = dataFromMasterReport.Item(dEST)
		' .Cells(ReportRow, repActCol).Value    = dataFromMasterReport.Item(dACT)
		' .Cells(ReportRow, repCheckCol).Value  = dataFromMasterReport.Item(dCheck)
		
	End With

End Sub

'initialize variables for columns'
Sub initVar()

     'Columns for the "Report Tab"'
	 repJobIdCol   = 1
	 repDateCol    = 2
	 repkWCol      = 3
	 repStatusCol  = 4
	 repOldNewCol  = 5
	 repPaidOutCol = 6
	 repCurValCol  = 7
	 repEstCol     = 8
	 repActCol     = 9
	 repCheckCol   = 10


	 'Columns for the "Master Report" Tab'
	 masJobIdCol     = 2
	 masDateCol      = 7
	 maskWCol        = 3
	 masStatusCol    = 4
	 masFinalCol     = 8
	 masCancelledCol = 19
	 masInstallCol   = 20

	 'Collection Keys'
	 dJOBID     = "jobID"
	 dKW        = "kW"
	 dSTATUS    = "Status"
	 dDATE      = "Date"
	 dFINAL     = "Final"
	 dINSTALL   = "Installed"
	 dCANCELLED = "Cancelled"

End Sub

'Checks which payout structure this account falls under
Sub check_structure(ByVal ReportRow, ByVal repDateCol, ByVal repOldNewCol)
    With Sheets("Report")
        If .Cells(ReportRow, repDateCol) < 41974 Then
            .Cells(ReportRow, repOldNewCol) = "Old"
            Call old_payout_structure
        Else
            .Cells(ReportRow, repOldNewCol) = "New"
            Call new_payout_structure
        End If
    End With
End Sub

'Sub for New Payout Structure
Sub new_payout_structure()



End Sub

'Sub for Old payout structure
Sub old_payout_structure()



End Sub
