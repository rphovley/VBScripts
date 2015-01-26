'RUN MODULE 8 TO REVERSE FIRST AND LAST NAME
'OF COLUMN A IN THE REPORT AND SORT REPORT
'BY createdOn
Sub newSales()
'RUN MODULE 8 TO REVERSE FIRST AND LAST NAME
'OF COLUMN A IN THE REPORT AND SORT REPORT BY
'createdOn

'CHECK OVERRIDEMONTH AND MONTHNUMBERPREV


    'report columns'
    Dim CustomerCol, JobIDCol, kWCol, StatusCol, SubStatusCol, theDateCol, repEmailCol As Integer

    CustomerCol = 1
    JobIDCol = 2
    kWCol = 3
    StatusCol = 4
    SubStatusCol = 5
    theDateCol = 7
    repEmailCol = 17

    'master report workbook'
    Dim masterReport As Workbook

    Dim i, j, x, startRow, histRow, payRow, startID As Integer
    Dim Master As Workbook
    Dim repRef, findRep, report As Worksheet
    Dim FilePath, FileName, histSheet, reportSheet, closedWonSheet, JobID, overrideMonth, _
    monthNumber, overrideYear, repEmail, repName As String
    Dim cancel, cancelled, Sales As String
    Dim isSale As Boolean

    cancel = "Cancel"
    cancelled = "Cancelled"
    Sales = "Sales"
    closedWonSheet = "report"
    Dim isCancelled As Boolean
    repName = ""

    'CHECK THESE
    overrideMonth = "December"
    overrideYear = "2014"
    monthNumberPrev = 12
    'CHECK THESE

    i = 2


    MsgBox ("Open Master Report")

    FilePath = Application.GetOpenFilename()
    FileName = convertToName(FilePath)

    If isWorkBookOpen(FilePath) Then
        Set masterReport = Workbooks(FileName)
    Else
        Set masterReport = Workbooks.Open(FilePath)
    End If


    Set Master = Workbooks(ThisWorkbook.Name)
    Set report = masterReport.Sheets("Current Data")
    Set repRef = Master.Sheets(overrideMonth & " " & overrideYear & " Map")
    Set findRep = Master.Sheets("RepsEmail")
    Set histSheet = Master.Sheets("Override Past")

    With Sheets("Payments")
        startID = Application.WorksheetFunction.Max(.Range("A:A")) + 1

        startRow = .Cells(1, 2).End(xlDown).Row + 1
        histRow = histSheet.Cells(1, 2).End(xlDown).Row + 1
        payRow = startRow

        Do Until IsEmpty(report.Cells(i, 1))
            x = 1
            isSale = False

                If Month(report.Cells(i, theDateCol).Value) = monthNumberPrev Then
                    If report.Cells(i, StatusCol).Value <> cancel And _
                        report.Cells(i, StatusCol).Value <> Sales And _
                        report.Cells(i, StatusCol).Value <> cancelled Then
                        If report.Cells(i, StatusCol).Value = "Permit" Then
                            isSale = isReady(report.Cells(i, SubStatusCol).Value)
                        Else
                            isSale = True
                        End If

                        If isSale Then

                        repEmail = report.Cells(i, repEmailCol)

                  On Error GoTo emailError:
                            repName = Application.WorksheetFunction.Index(findRep.Range("G:G"), _
                                Application.WorksheetFunction.Match(repEmail, findRep.Range("B:B"), 0))

                            If repName <> "" Then
                                'Enter Data in Override past
                                x = 0
                                'Rep
                                histSheet.Cells(histRow, 1).Value = repName
                                'Date
                                histSheet.Cells(histRow, 2).Value = overrideMonth
                                'JobID
                                histSheet.Cells(histRow, 3).Value = report.Cells(i, JobIDCol).Value
                                'Customer
                                histSheet.Cells(histRow, 4).Value = report.Cells(i, CustomerCol).Value
                                'kW
                                histSheet.Cells(histRow, 5).Value = report.Cells(i, kWCol).Value
                                'Reason
                                histSheet.Cells(histRow, 6).Value = "New Sale"
                                'Entry Date
                                histSheet.Cells(histRow, 7).Value = overrideMonth

                                On Error Resume Next
                                x = WorksheetFunction.Match(repName, repRef.Range("A:A"), 0)


                                'Do Until IsEmpty(repRef.Cells(x, 1))
                                    'Check if the cell row is for the right rep
                                    If x <> 0 Then
                                        'Set repID to the ID of the cell referenced
                                        repID = repRef.Cells(x, 2).Value
                                        rowLength = repRef.Cells(1, 1).End(xlToRight).Row
                                        'loop through column to get all overrides

                                        For y = 3 To repRef.Cells(1, 1).End(xlToRight).Column

                                            If repRef.Cells(x, y).Value <> "" And repRef.Cells(x, y + 1).Value <> vbNull Then

                                                'OverrideID
                                                .Cells(payRow, 1).Value = startID
                                                'OverrideRep
                                                .Cells(payRow, 2).Value = repRef.Cells(x, y + 1).Value
                                                'OverrideRepID
                                                .Cells(payRow, 3).Value = repRef.Cells(x, y + 2).Value
                                                'Rep
                                                .Cells(payRow, 4).Value = repName

                                                'RepID
                                                .Cells(payRow, 5).Value = repRef.Cells(x, 2).Value
                                                'Customer
                                                .Cells(payRow, 6).Value = report.Cells(i, CustomerCol).Value

                                                'JobID
                                                .Cells(payRow, 7).Value = report.Cells(i, JobIDCol).Value

                                                'Date
                                                .Cells(payRow, 8).Value = overrideMonth + " " + overrideYear

                                                'Date of Override
                                                .Cells(payRow, 9).Value = overrideMonth + " " + overrideYear

                                                'Reason
                                                .Cells(payRow, 10).Value = "New Sale"
                                                'Status
                                                .Cells(payRow, 11).Value = report.Cells(i, StatusCol).Value
                                                'SubStatus
                                                .Cells(payRow, 12).Value = report.Cells(i, SubStatusCol).Value
                                                'OverrideType
                                                .Cells(payRow, 13).Value = repRef.Cells(x, y).Value
                                                'OverrideRate
                                                .Cells(payRow, 14).Value = repRef.Cells(x, y + 3).Value
                                                'kW
                                                .Cells(payRow, 15).Value = report.Cells(i, kWCol).Value

                                                'Amount
                                                .Cells(payRow, 16).Value = (.Cells(payRow, 14).Value * .Cells(payRow, 15).Value) / 2

                                                payRow = payRow + 1

                                                startID = startID + 1
                                            End If
                                           y = y + 3
                                        Next y

                                    End If

                                'Loop
                                histRow = histRow + 1
                            End If
                      End If
                End If
            End If
        i = i + 1
        Loop

    End With

emailError:
    report.Rows(i).Interior.Color = vbRed
    Resume Next

End Sub


