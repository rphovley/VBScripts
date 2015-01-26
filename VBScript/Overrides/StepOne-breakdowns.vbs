Private Sub cmdRun_Click()

'Make main set of variables
Dim Rep As String
Dim Email As String

Dim subRep As String
Dim customer As String
Dim kW As Variant
Dim Rate As String
Dim Total As Currency
Dim Reason As String
Dim JobID As String
Dim iStatus As String
Dim SubStatus As String
Dim OverrideType As String
Dim ReportDate As Date

ReportDate = InputBox("Input this friday's date as mm/dd/yyyy")

'Make row counter variables
Dim masterrow As Integer
Dim reportrow As Integer
Dim reprow As Integer

Dim repCol, customerCol, kWCol, rateCol, totalCol, reasonCol, jobIDCol, statusCol, subStatusCol, overTypeCol As Integer

repCol = 2
customerCol = 3
kWCol = 4
rateCol = 5
totalCol = 6
reasonCol = 7
jobIDCol = 8
statusCol = 9
subStatusCol = 10
overTypeCol = 11

'must be outside of all loops so that it doesn't reset
reprow = 1

'Main part of code that loops through the reps
Do
'Set starting row for the row counters for each loop/rep
reportrow = 5
masterrow = 2
    'Sets the Rep as the new rep to be done
    Rep = Sheets("Reps").Cells(reprow, 1)
    'Creates a new tab with the same name as the rep
    Worksheets.Add(, Worksheets(Worksheets.Count)).Name = Rep
    'Formats the new spreadsheet
    Worksheets(Rep).Cells(1, 4) = "Name:"
    Worksheets(Rep).Cells(1, 5) = Rep
    Worksheets(Rep).Cells(2, 4) = "Date:"
    Worksheets(Rep).Cells(2, 5) = ReportDate
    Worksheets(Rep).Cells(4, repCol) = "Rep"
    Worksheets(Rep).Cells(4, customerCol) = "Customer"
    Worksheets(Rep).Cells(4, kWCol) = "kW"
    Worksheets(Rep).Cells(4, rateCol) = "Rate/kW"
    Worksheets(Rep).Cells(4, totalCol) = "Total"
    Worksheets(Rep).Cells(4, reasonCol) = "Type"
    Worksheets(Rep).Cells(4, jobIDCol) = "Job ID"
    Worksheets(Rep).Cells(4, statusCol) = "Status"
    Worksheets(Rep).Cells(4, subStatusCol) = "SubStatus"
    Worksheets(Rep).Cells(4, overTypeCol) = "OverrideType"

    'Formats the Name and Date delineators
    With Worksheets(Rep).Range(Sheets(Rep).Cells(1, 3), Sheets(Rep).Cells(2, 3))
        .HorizontalAlignment = xlRight
        .Font.Bold = True
    End With


    'Formats the column headers
    With Worksheets(Rep).Range("B4:K4")
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With


    'Loops through each account before moving on
    Do
        If Sheets("Master Sheet").Cells(masterrow, 2) = Rep Then

        'Gives values to the variables to be put into the created worksheet
        With Sheets("Master Sheet")
          subRep = .Cells(masterrow, 4)
          customer = .Cells(masterrow, 6)
          kW = .Cells(masterrow, 15)
          Rate = .Cells(masterrow, 14)
          Total = .Cells(masterrow, 16)
          Reason = .Cells(masterrow, 10)
          JobID = .Cells(masterrow, 7)
          iStatus = .Cells(masterrow, 11)
          SubStatus = .Cells(masterrow, 12)
          OverrideType = .Cells(masterrow, 13)
        End With

        'Inputs data into the rep's report
        Sheets(Rep).Cells(reportrow, repCol) = subRep
        Sheets(Rep).Cells(reportrow, customerCol) = customer
        Sheets(Rep).Cells(reportrow, kWCol) = kW
        Sheets(Rep).Cells(reportrow, rateCol) = Rate
        Sheets(Rep).Cells(reportrow, totalCol) = Total
        Sheets(Rep).Cells(reportrow, reasonCol) = Reason
        Sheets(Rep).Cells(reportrow, jobIDCol) = JobID
        Sheets(Rep).Cells(reportrow, statusCol) = iStatus
        Sheets(Rep).Cells(reportrow, subStatusCol) = SubStatus
        Sheets(Rep).Cells(reportrow, overTypeCol) = OverrideType

        'Report row counter only moves if data was copied into the rep's spreadsheet
        reportrow = reportrow + 1
        End If
        masterrow = masterrow + 1
    Loop Until Sheets("Master Sheet").Cells(masterrow, 1) = ""
    'Sums the totals for each customer into a grand total for the rep
    Worksheets(Rep).Cells(reportrow + 1, 5) = "Total:"
    Worksheets(Rep).Cells(reportrow + 1, 6).Formula = "=Sum(" & Range(Cells(5, 6), Cells(reportrow, 6)).Address() & ")"
    'Formats the grand total cells
    With Worksheets(Rep).Cells(reportrow + 1, 5)
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    'Places border around the grand total
    With Worksheets(Rep).Range(Sheets(Rep).Cells(reportrow + 1, 5), Sheets(Rep).Cells(reportrow + 1, 6))
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With

    'Adjusts column width of the new tab
    With Worksheets(Rep).Range("B1:F1")
        .EntireColumn.AutoFit
    End With

    'Moves the rep row counter to the next rep
    reprow = reprow + 1

    'Creates a workbook for each rep
    'ThisWorkbook.Sheets(Rep).Copy
    'ActiveWorkbook.SaveAs ("C:\users\Rodney\desktop\" & "Payroll Breakdown\" & Rep & ".xlsx")
    'ActiveWorkbook.Close

Loop Until Sheets("Reps").Cells(reprow, 1) = ""

End Sub
