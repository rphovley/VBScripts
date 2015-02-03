	'Columns for the "Report Tab"'
	Dim repJobIdCol, repDateCol, repkWCol, repStatusCol, _
	 repOldNewCol, repEstCol, repActCol, repCheckCol As Integer

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

	'dim the row vars for both tabs'
	Dim MasterReportRow, ReportRow As Integer

	MasterReportRow = 2
	ReportRow       = 2

	'used to pass information back and forth from functions'
	Dim dataFromMasterReport As New Collection

	'Loop through the Master Report'
	Do Until isEmpty(Sheets("Master Report").Cells(MasterReportRow, 1).Value)
	
		'Collect Data from Master Report and Determine what should be paid out to us in the Master Report'
		dataFromMasterReport = determinePayout(dataFromMasterReport, MasterReportRow)
		'print out what should be paid out in the Report Tab'
	 	printData dataFromMasterReport
	Loop


End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''Supporting Subs and Functions''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Determine what should be paid out to us in the Master Report'
Function determinePayout(ByRef dataFromMasterReport As Collection, ByVal MasterReportRow As Integer)
	Dim jobID, kW, Status, date

	'Collection data from Master Report'
	dataFromMasterReport = setCollection(dataFromMasterReport, MasterReportRow)

	'We make some decision based on what we find in the report'

	'This is returning the collection to the calling sub'
	
	determinePayout = dataFromMasterReport
End Function

'Set Collection Values for the data from the Master Report'
Function setCollection(ByRef dataFromMasterReport As Collection, ByVal MasterReportRow As Integer)
	Dim jobID, kW, Status, date

	With Sheets("Master Report")
		dataFromMasterReport.Add .Cells(MasterReportRow, masJobIdCol) dJOBID
	    dataFromMasterReport.Add .Cells(MasterReportRow, maskWCol), dKW
	    dataFromMasterReport.Add .Cells(MasterReportRow, masStatusCol), dSTATUS
	    dataFromMasterReport.Add .Cells(MasterReportRow, masDateCol), dDATE
	    dataFromMasterReport.Add .Cells(MasterReportRow, masFinalCol), dFINAL
	    dataFromMasterReport.Add .Cells(MasterReportRow, masCancelledCol), dCANCELLED
	    dataFromMasterReport.Add .Cells(MasterReportRow, masInstallCol), dINSTALL
	End With

	'this is returning the collection from the calling function'
	setCollection = dataFromMasterReport

End Function

'Sub to print out data gathered into the Report Tab'
Sub printData(ByRef dataFromMasterReport)
	
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
	 repJobIdCol  = 1
	 repDateCol   = 2
	 repkWCol     = 3
	 repStatusCol = 4
	 repOldNewCol = 5
	 repEstCol    = 6
	 repActCol    = 7
	 repCheckCol  = 8


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