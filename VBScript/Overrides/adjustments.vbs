Dim solarReport, Master, closedWon As Workbook
    Dim map, report, repRef, findRep, closedRep, histSheet As Worksheet
Dim overrideMonth, overrideYear As String
Dim bottomRow As Integer
Dim isFirst As Boolean
'report columns'
    Dim CustomerCol, JobIDCol, kWCol, StatusCol, SubStatusCol, theDateCol, repEmailCol As Integer
'SORT BY JOB ID THEN OVERRIDE ID BEFORE YOU RUN THIS
Sub newAdjustments()
'SORT BY JOB ID THEN OVERRIDE ID BEFORE YOU RUN THIS

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''INIT VARIABLES'''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    CustomerCol = 1
    JobIDCol = 2
    kWCol = 3
    StatusCol = 4
    SubStatusCol = 5
    theDateCol = 7
    repEmailCol = 17

    'master report workbook'
    Dim masterReport As Workbook
    Dim i, j, x, y, histRow, jobRow, overRepID, rowLength As Integer

    Dim FilePath, FileName, reportSheet, closedWonSheet, JobID, repName As String
    Dim isCancelled, isCurrentlyCancelled, paymentsIdFound, isJobBackend, isAlreadyBackend, jobIDfound, isSale As Boolean
    Dim dictRows As New Collection


    overrideMonth = "x"
    overrideYear = "0"

    i = 1

    'MsgBox ("Open Most Recent Solar City Report")

    FilePath = Application.GetOpenFilename()
    FileName = convertToName(FilePath)

    If isWorkBookOpen(FilePath) Then
        Set masterReport = Workbooks(FileName)
    Else
        Set masterReport = Workbooks.Open(FilePath)
    End If

    Set Master = Workbooks(ThisWorkbook.Name)
    Set report = masterReport.Sheets("Current Data")
    Set findRep = Master.Sheets("RepsEmail")
    Set histSheet = Master.Sheets("Override Past")

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''LOGIC'''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
With Master.Sheets("Payments")

    bottomRow = .Cells(1, 1).End(xlDown).Row + 1
    'loop through each record in the master report
    Do Until IsEmpty(report.Cells(i, 1))
        Dim sum, kWsum, checkID As Double
        jobIDfound = True
        paymentsIdFound = True
        isAlreadyBackend = False

        FileName = report.Cells(i, JobIDCol).Value

        'find first row associated with jobID if it is not found, it is a new sale
        On Error GoTo jobIdNotFound:
            JobID = report.Cells(i, JobIDCol).Value
            jobRow = Application.WorksheetFunction.Match(report.Cells(i, JobIDCol).Value, Master.Sheets("Override Past").Range("C:C"), 0)

        On Error GoTo paymentIdNotFound:
            jobRow = Application.WorksheetFunction.Match(report.Cells(i, JobIDCol).Value, .Range("G:G"), 0)

        'find overrideID for first related override

        y = jobRow
        JobID = .Cells(y, 7).Value
        checkID = .Cells(y, 3).Value

        isFirst = True

        If jobIDfound Then
            Dim checkStatus As String
            Dim isItCancelled As Boolean
            checkStatus = report.Cells(i, StatusCol).Value
            If checkStatus <> "Cancelled" Then
                isItCancelled = isJobCancelled(report.Cells(i, SubStatusCol).Value)
            Else
                isItCancelled = True
            End If
            'Loop through all JobIDs that equal the jobID on the report
            Set repRef = Master.Sheets(overrideMonth + " " + overrideYear + " Map")

                JobID = report.Cells(i, JobIDCol).Value
            'Was the id found in the payments tab, if it wasn't it is a backend for a rep that didn't have anyone for overrides before
            If paymentsIdFound Then
                isAlreadyBackend = False
                Do Until .Cells(y, 7).Value <> report.Cells(i, JobIDCol).Value
                    On Error GoTo wtf:
                    overRepID = Application.WorksheetFunction.Index(.Range("C:C"), Application.WorksheetFunction.Match(.Cells(y, 2).Value, .Range("B:B"), 0))
                    On Error Resume Next:
                    overID = Application.WorksheetFunction.Index(.Range("A:A"), y)
                    sum = 0
                    kWsum = 0
                    isCurrentlyCancelled = False

                    x = y
                    'Loop through the individual overrides associated with this jobID
                    Do Until .Cells(x, 1).Value <> overID
                        sum = sum + .Cells(y + (x - y), 16).Value
                        kWsum = kWsum + .Cells(y + (x - y), 15).Value
                        If LCase(.Cells(x, 10).Value) = "cancelled" Then
                            isCurrentlyCancelled = True
                        End If
                        If .Cells(x, 8).Value = "Backend" Then
                            isAlreadyBackend = True
                        End If
                        x = x + 1
                    Loop

                    dictRows.Add i, "solarRow"
                    dictRows.Add y, "jobRow"
                    dictRows.Add sum, "sum"
                    dictRows.Add kWsum, "kWsum"
                    dictRows.Add bottomRow, "bottomRow"


                    'was the jobID already cancelled?
                    If isCurrentlyCancelled Then

                    'was not cancelled
                    Else

                        'If it is a backend
                        If isAlreadyBackend = False Then

                            'if it is a Backend
                            If isBackend_New(report.Cells(i, StatusCol).Value, report.Cells(i, SubStatusCol).Value) Then

                                    'record in historical jobs
                                    isFirst = toHist(i, .Cells(y, 4).Value, "New Sale", isFirst)
                                    'OverrideID
                                    .Cells(bottomRow, 1).Value = overID
                                    'OverrideRep
                                    .Cells(bottomRow, 2).Value = .Cells(y, 2).Value
                                    'OverrideRepID
                                    .Cells(bottomRow, 3).Value = overRepID
                                    'Rep
                                    .Cells(bottomRow, 4).Value = .Cells(y, 4).Value
                                    'RepID
                                    .Cells(bottomRow, 5).Value = .Cells(y, 5).Value
                                    'Customer
                                    .Cells(bottomRow, 6).Value = .Cells(y, 6).Value
                                    'JobID
                                    .Cells(bottomRow, 7).Value = report.Cells(i, JobIDCol).Value
                                    'Date
                                    .Cells(bottomRow, 8).Value = "Backend"
                                    'Entry Date
                                    .Cells(bottomRow, 9).Value = overrideMonth & " " & overrideYear
                                    'Reason
                                    .Cells(bottomRow, 10).Value = "Backend"
                                    'Status
                                    .Cells(bottomRow, 11).Value = report.Cells(i, StatusCol).Value
                                    'SubStatus
                                    .Cells(bottomRow, 12).Value = report.Cells(i, SubStatusCol).Value
                                    'OverrideType
                                    .Cells(bottomRow, 13).Value = .Cells(y, 13).Value
                                    'OverrideRate
                                    .Cells(bottomRow, 14).Value = .Cells(y, 14).Value
                                    'kW AND Amount
                                    Dim newTotal As Double
                                    'has kW changed since last override? If it has, the total amount due needs to be recalculated
                                    If report.Cells(i, kWCol).Value <> kWsum Then
                                        'kW
                                        .Cells(bottomRow, 15).Value = report.Cells(i, kWCol).Value

                                        newTotal = .Cells(bottomRow, 14).Value * .Cells(bottomRow, 15).Value
                                        'BackendAmount
                                        .Cells(bottomRow, 16).Value = newTotal - .Cells(y, 16).Value
                                    Else
                                        'kW
                                        .Cells(bottomRow, 15).Value = kWsum
                                        'BackendAmount
                                        newTotal = .Cells(bottomRow, 14).Value * .Cells(bottomRow, 15).Value
                                        .Cells(bottomRow, 16).Value = newTotal - sum
                                    End If

                                    bottomRow = bottomRow + 1


                            'If it is cancelled
                             ElseIf isItCancelled = True Then
                                'record in historical jobs
                                isFirst = toHist(i, .Cells(y, 4).Value, "New Sale", isFirst)
                                'OverrideID
                                .Cells(bottomRow, 1).Value = overID
                                'OverrideRep
                                .Cells(bottomRow, 2).Value = .Cells(y, 2).Value
                                'OverrideRepID
                                .Cells(bottomRow, 3).Value = overRepID
                                'Rep
                                .Cells(bottomRow, 4).Value = .Cells(y, 4).Value
                                'RepID
                                .Cells(bottomRow, 5).Value = .Cells(y, 5).Value
                                'Customer
                                .Cells(bottomRow, 6).Value = .Cells(y, 6).Value
                                'JobID
                                .Cells(bottomRow, 7).Value = report.Cells(i, JobIDCol).Value
                                'Date
                                .Cells(bottomRow, 8).Value = .Cells(y, 8).Value
                                'Entry Date
                                .Cells(bottomRow, 9).Value = overrideMonth & " " & overrideYear
                                'Reason
                                .Cells(bottomRow, 10).Value = "Cancelled"
                                'Status
                                .Cells(bottomRow, 11).Value = report.Cells(i, StatusCol).Value
                                'SubStatus
                                .Cells(bottomRow, 12).Value = report.Cells(i, SubStatusCol).Value
                                'OverrideType
                                .Cells(bottomRow, 13).Value = .Cells(y, 13).Value
                                'OverrideRate
                                .Cells(bottomRow, 14).Value = .Cells(y, 14).Value
                                'kW
                                .Cells(bottomRow, 15).Value = kWsum
                                'Amount
                                .Cells(bottomRow, 16).Value = -sum

                                bottomRow = bottomRow + 1


                            'was not cancelled and is not a backend
                            Else
                                Dim testKw, testKw2 As String
                                testKw = report.Cells(i, kWCol).Value
                                testKw2 = report.Cells(i, kWCol).Value
                                'does the kW of the override equal the report kilowatt?
                                If (report.Cells(i, kWCol).Value - kWsum) > 0.1 Or (report.Cells(i, kWCol).Value - kWsum) < -0.1 Then
                                    'adjustment equal to the rate times the difference between new and old kW

                                    'OverrideID
                                    .Cells(bottomRow, 1).Value = overID
                                    'OverrideRep
                                    .Cells(bottomRow, 2).Value = .Cells(y, 2).Value
                                    'OverrideRepID
                                    .Cells(bottomRow, 3).Value = .Cells(y, 3).Value
                                    'Rep
                                    .Cells(bottomRow, 4).Value = .Cells(y, 4).Value
                                    'RepID
                                    .Cells(bottomRow, 5).Value = .Cells(y, 5).Value
                                    'Customer
                                    .Cells(bottomRow, 6).Value = .Cells(y, 6).Value
                                    'JobID
                                    .Cells(bottomRow, 7).Value = report.Cells(i, JobIDCol).Value
                                    'Date
                                    .Cells(bottomRow, 8).Value = .Cells(y, 8).Value
                                    'Entry Date
                                    .Cells(bottomRow, 9).Value = overrideMonth & " " & overrideYear
                                    'Reason
                                    'ignore rounding
                                    If report.Cells(i, kWCol).Value - kWsum > 0 Then
                                        .Cells(bottomRow, 10).Value = "Upsize"
                                        'record in historical jobs
                                        isFirst = toHist(i, .Cells(y, 4).Value, "Upsize", isFirst)
                                    Else
                                        .Cells(bottomRow, 10).Value = "Downsize"
                                        'record in historical jobs
                                        isFirst = toHist(i, .Cells(y, 4).Value, "Upsize", isFirst)
                                    End If
                                    Dim roundChecks As Double
                                    roundChecks = Round(report.Cells(i, kWCol).Value, 2)
                                    'Status
                                    .Cells(bottomRow, 11).Value = report.Cells(i, StatusCol).Value
                                    'SubStatus
                                    .Cells(bottomRow, 12).Value = report.Cells(i, SubStatusCol).Value
                                    'OverrideType
                                    .Cells(bottomRow, 13).Value = .Cells(y, 13).Value
                                    'OverrideRate
                                    .Cells(bottomRow, 14).Value = .Cells(y, 14).Value
                                    'kW
                                    .Cells(bottomRow, 15).Value = (report.Cells(i, kWCol).Value - kWsum)
                                    'Amount
                                    testKw = .Cells(y, 14).Value
                                    testKw = report.Cells(i, kWCol).Value
                                    testKw2 = (.Cells(y, 14).Value * (report.Cells(i, kWCol).Value - kWsum)) / 2
                                    .Cells(bottomRow, 16).Value = (.Cells(y, 14).Value * (report.Cells(i, kWCol).Value - kWsum)) / 2

                                    bottomRow = bottomRow + 1
                                End If
                            End If
                        End If

                    End If
                    'set overrideID equal to the next override associated with the jobID
                    overID = .Cells(x, 1).Value
                    'isFirst = False
                    dictRows.Remove "solarRow"
                    dictRows.Remove "jobRow"
                    dictRows.Remove "sum"
                    dictRows.Remove "kWsum"
                    dictRows.Remove "bottomRow"
                y = x

                Loop

            'Anything that was found in the historical payments tab that hasn't been paid out on, that is now needing payment (backends)
            Else
                'repName found in Override Past
                repName = Application.WorksheetFunction.Index(Master.Sheets("Override Past").Range("A:A"), Application.WorksheetFunction.Match(report.Cells(i, JobIDCol).Value, Master.Sheets("Override Past").Range("C:C"), 0))

                'if it is a Backend
                If isBackend_New(report.Cells(i, StatusCol).Value, report.Cells(i, SubStatusCol).Value) Then
                    x = 1
                    overID = Application.WorksheetFunction.Max(.Range("A:A")) + 1
                    If Month(report.Cells(i, theDateCol).Value) > 5 And Month(report.Cells(i, theDateCol).Value) < 11 Then
                        Set repRef = Master.Sheets(MonthName(Month(report.Cells(i, theDateCol).Value), False) & " " & Year(report.Cells(i, theDateCol).Value) & " Map")
                    Else
                        Set repRef = Master.Sheets("May 2014 Map")
                    End If
                    'record in historical jobs
                    isFirst = toHist(i, repName, "Upsize", isFirst)
                 Do Until IsEmpty(repRef.Cells(x, 1))
                            'Check if the cell row is for the right rep
                            If repRef.Cells(x, 1).Value = repName Then
                                'Set repID to the ID of the cell referenced
                                repID = repRef.Cells(x, 2).Value
                                rowLength = repRef.Cells(1, 1).End(xlToRight).Row
                                'loop through column to get all overrides
                                For y = 3 To repRef.Cells(1, 1).End(xlToRight).Column
                                    If repRef.Cells(x, y).Value <> "" And repRef.Cells(x, y + 1).Value <> vbNull Then
                                        'OverrideID
                                        .Cells(bottomRow, 1).Value = overID
                                        'OverrideRep
                                        .Cells(bottomRow, 2).Value = repRef.Cells(x, y + 1).Value
                                        'OverrideRepID
                                        .Cells(bottomRow, 3).Value = repRef.Cells(x, y + 2).Value
                                        'Rep
                                        .Cells(bottomRow, 4).Value = repName
                                        'RepID
                                        .Cells(bottomRow, 5).Value = repRef.Cells(x, 2).Value
                                        'Customer
                                        .Cells(bottomRow, 6).Value = report.Cells(i, CustomerCol).Value
                                        'JobID
                                        .Cells(bottomRow, 7).Value = report.Cells(i, JobIDCol).Value
                                        'Date
                                        .Cells(bottomRow, 8).Value = "Backend"
                                        'Entry Date
                                        .Cells(bottomRow, 9).Value = overrideMonth & " " & overrideYear
                                        'Reason
                                        .Cells(bottomRow, 10).Value = "Backend"
                                        'Status
                                        .Cells(bottomRow, 11).Value = report.Cells(i, StatusCol).Value
                                        'SubStatus
                                        .Cells(bottomRow, 12).Value = report.Cells(i, SubStatusCol).Value
                                        'OverrideType
                                        .Cells(bottomRow, 13).Value = repRef.Cells(x, y).Value
                                        'OverrideRate
                                        .Cells(bottomRow, 14).Value = repRef.Cells(x, y + 3).Value
                                        'kW
                                        .Cells(bottomRow, 15).Value = report.Cells(i, kWCol).Value
                                        'Amount
                                        .Cells(bottomRow, 16).Value = (.Cells(bottomRow, 14).Value * .Cells(bottomRow, 15).Value) / 2

                                        bottomRow = bottomRow + 1
                                        overID = overID + 1
                                     End If
                                   y = y + 3
                                Next y

                            End If
                        x = x + 1

                        Loop
                End If

            End If

    'IF is New Sale (and possibly all the way to backend)
    Else
        isJobBackend = False
        x = 1
        Dim testString As String

        testString = MonthName(Month(report.Cells(i, theDateCol).Value), False) & " " & Year(report.Cells(i, theDateCol).Value) & " Map"
        If Month(report.Cells(i, theDateCol).Value) > 5 And Month(report.Cells(i, theDateCol).Value) < 11 Then
            Set repRef = Master.Sheets(MonthName(Month(report.Cells(i, theDateCol).Value), False) & " " & Year(report.Cells(i, theDateCol).Value) & " Map")
        Else
            Set repRef = Master.Sheets("May 2014 Map")
        End If

        overID = Application.WorksheetFunction.Max(.Range("A:A")) + 1
        isSale = False

        Dim cancel, cancelled, Sales As String
        cancel = "Cancel"
        cancelled = "Cancelled"
        Sales = "Sales"
        repName = ""
        histRow = Master.Sheeets("Override Past").Cells(1, 2).End(xlDown).Row + 1

        If report.Cells(i, StatusCol).Value <> cancel And report.Cells(i, StatusCol).Value <> Sales And report.Cells(i, StatusCol).Value <> cancelled Then
                    If report.Cells(i, StatusCol).Value = "Permit" Then
                        isSale = isReady(report.Cells(i, SubStatusCol).Value)
                    Else
                        isJobBackend = isBackend_New(report.Cells(i, StatusCol).Value, report.Cells(i, SubStatusCol).Value)
                        isSale = True
                    End If

                    If isSale Then

                    repEmail = ""
                    repEmail = report.Cells(i, repEmailCol).Value

                    On Error GoTo emailError:
                        repName = Application.WorksheetFunction.Index(findRep.Range("G:G"), Application.WorksheetFunction.Match(repEmail, findRep.Range("B:B"), 0))

                        If repName <> "" And repEmail <> "" Then

                            histRow = histRow + 1
                            Do Until IsEmpty(repRef.Cells(x, 1))
                                'Check if the cell row is for the right rep
                                If repRef.Cells(x, 1).Value = repName Then
                                    'Set repID to the ID of the cell referenced
                                    repID = repRef.Cells(x, 2).Value
                                    rowLength = repRef.Cells(1, 1).End(xlToRight).Row
                                    'loop through column to get all overrides
                                    For y = 3 To repRef.Cells(1, 1).End(xlToRight).Column

                                        If repRef.Cells(x, y).Value <> "" And repRef.Cells(x, y + 1).Value <> vbNull Then

                                            'OverrideID
                                            .Cells(bottomRow, 1).Value = overID
                                            'OverrideRep
                                            .Cells(bottomRow, 2).Value = repRef.Cells(x, y + 1).Value
                                            'OverrideRepID
                                            .Cells(bottomRow, 3).Value = repRef.Cells(x, y + 2).Value
                                            'Rep
                                            .Cells(bottomRow, 4).Value = repName
                                            'RepID
                                            .Cells(bottomRow, 5).Value = repRef.Cells(x, 2).Value
                                            'Customer
                                            .Cells(bottomRow, 6).Value = report.Cells(i, CustomerCol).Value
                                            'JobID
                                            .Cells(bottomRow, 7).Value = report.Cells(i, JobIDCol).Value
                                            'Date
                                            If isJobBackend Then
                                                .Cells(bottomRow, 8).Value = "Backend"
                                            Else
                                                .Cells(bottomRow, 8).Value = MonthName(Month(report.Cells(i, theDateCol).Value), False) & " " & overrideYear

                                            End If
                                            'Date of Override
                                            .Cells(bottomRow, 9).Value = overrideMonth & " " & overrideYear
                                            'Reason
                                             If isJobBackend Then
                                                .Cells(bottomRow, 10).Value = "Backend"
                                                'record in historical jobs
                                                isFirst = toHist(i, repName, "Upsize", isFirst)
                                             Else
                                                .Cells(bottomRow, 10).Value = "New Sale"
                                                'record in historical jobs
                                                isFirst = toHist(i, repName, "Upsize", isFirst)
                                             End If
                                             'Status
                                            .Cells(bottomRow, 11).Value = report.Cells(i, StatusCol).Value
                                            'SubStatus
                                            .Cells(bottomRow, 12).Value = report.Cells(i, SubStatusCol).Value
                                            'OverrideType
                                            .Cells(bottomRow, 13).Value = repRef.Cells(x, y).Value
                                            'OverrideRate
                                            .Cells(bottomRow, 14).Value = repRef.Cells(x, y + 3).Value
                                            'kW
                                            .Cells(bottomRow, 15).Value = report.Cells(i, kWCol).Value
                                            'Amount
                                            .Cells(bottomRow, 16).Value = (.Cells(bottomRow, 14).Value * .Cells(bottomRow, 15).Value) / 2

                                            bottomRow = bottomRow + 1
                                            overID = overID + 1
                                        End If
                                       y = y + 3
                                    Next y

                                End If
                            x = x + 1
                            Loop
                        End If
                  End If
            End If
    End If

    i = i + 1
    Loop
End With

jobIdNotFound:
    jobIDfound = False
    Resume Next
paymentIdNotFound:
    paymentsIdFound = False
    Resume Next
wtf:
    Resume Next
emailError:
    report.Rows(i).Interior.Color = vbRed
    Resume Next
End Sub

Function toHist(ByVal i As Integer, ByVal repName As String, ByVal Reason As String, ByVal isFirst As Boolean) As Boolean
          Dim histRow As Integer

        If isFirst Then
           histRow = histSheet.Cells(1, 1).End(xlDown).Row + 1
              'Rep
           histSheet.Cells(histRow, 1).Value = repName
'              Date
           histSheet.Cells(histRow, 2).Value = overrideMonth & " " & overrideYear
              'JobID
          histSheet.Cells(histRow, 3).Value = report.Cells(i, JobIDCol).Value
              'Customer
            histSheet.Cells(histRow, 4).Value = report.Cells(i, CustomerCol).Value
              'kW
           histSheet.Cells(histRow, 5).Value = report.Cells(i, kWCol).Value
              'Reason
            histSheet.Cells(histRow, 6).Value = Reason
              'Entry Date
            histSheet.Cells(histRow, 7).Value = overrideMonth & " " & overrideYear
        End If
            toHist = False
End Function





