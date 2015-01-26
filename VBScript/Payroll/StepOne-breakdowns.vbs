Private Sub cmdRun_Click()

'Make main set of variables
Dim Rep As String
Dim Email As String
Dim customer As String
Dim kW As Variant
Dim Total As Currency
Dim PaymentType As String
Dim JobID As String
Dim ReportDate As Date

ReportDate = InputBox("Input this friday's date as mm/dd/yyyy")

'Make row counter variables
Dim masterrow As Integer
Dim reportrow As Integer
Dim reprow As Integer

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
    Worksheets(Rep).Cells(1, 3) = "Name:"
    Worksheets(Rep).Cells(1, 4) = Rep
    Worksheets(Rep).Cells(2, 3) = "Date:"
    Worksheets(Rep).Cells(2, 4) = ReportDate
    Worksheets(Rep).Cells(4, 2) = "Customer"
    Worksheets(Rep).Cells(4, 3) = "kW"
    Worksheets(Rep).Cells(4, 4) = "Total"
    Worksheets(Rep).Cells(4, 5) = "Type"
    Worksheets(Rep).Cells(4, 6) = "Job ID"

    'Formats the Name and Date delineators
    With Worksheets(Rep).Range(Sheets(Rep).Cells(1, 3), Sheets(Rep).Cells(2, 3))
        .HorizontalAlignment = xlRight
        .Font.Bold = True
    End With


    'Formats the column headers
    With Worksheets(Rep).Range("B4:F4")
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
    End With


    'Loops through each account before moving on
    Do
        If Sheets("Master Sheet").Cells(masterrow, 1) = Rep Then

    'Gives values to the variables to be put into the created worksheet
        customer = Cells(masterrow, 2)
        kW = Cells(masterrow, 3)
        Total = Cells(masterrow, 4)
        PaymentType = Cells(masterrow, 5)
        JobID = Cells(masterrow, 6)

        'Inputs data into the rep's report
        Sheets(Rep).Cells(reportrow, 2) = customer
        Sheets(Rep).Cells(reportrow, 3) = kW
        Sheets(Rep).Cells(reportrow, 4) = Total
        Sheets(Rep).Cells(reportrow, 5) = PaymentType
        Sheets(Rep).Cells(reportrow, 6) = JobID

        'Report row counter only moves if data was copied into the rep's spreadsheet
        reportrow = reportrow + 1
        End If
        masterrow = masterrow + 1
    Loop Until Sheets("Master Sheet").Cells(masterrow, 1) = ""
    'Sums the totals for each customer into a grand total for the rep
    Worksheets(Rep).Cells(reportrow + 1, 3) = "Total:"
    Worksheets(Rep).Cells(reportrow + 1, 4).Formula = "=Sum(" & Range(Cells(5, 4), Cells(reportrow, 4)).Address() & ")"
    'Formats the grand total cells
    With Worksheets(Rep).Cells(reportrow + 1, 3)
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(0, 0, 0)
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    'Places border around the grand total
    With Worksheets(Rep).Range(Sheets(Rep).Cells(reportrow + 1, 3), Sheets(Rep).Cells(reportrow + 1, 4))
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
