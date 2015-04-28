Private Sub Override_History()

''''''''''''''''''TURN OFF SCREEN UPDATING''''''''''''''''''''''
Application.ScreenUpdating = False

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
Dim overrideType As String
Dim ReportDate As Date
Dim previous_month As String
Dim grand_total As Currency

ReportDate = InputBox("Input this friday's date as mm/dd/yyyy")

'Make row counter variables
Dim masterrow As Long
Dim reportrow As Integer
Dim repRow As Integer

'Make column counters
Dim repCol, customerCol, kWCol, rateCol, totalCol, reasonCol, jobIDCol, statusCol, subStatusCol, overTypeCol, month_col As Integer

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
month_col = 9



'must be outside of all loops so that it doesn't reset
repRow = 2

'Main part of code that loops through the reps
Do
'Set starting row for the row counters for each loop/rep
reportrow = 4
masterrow = 2
grand_total = 0

previous_month = UCase(Format(DateAdd("m", -1, ReportDate), "MMMM-YYYY"))


    'Sets the Rep as the new rep to be done
    Rep = Sheets("Reps").Cells(repRow, 1)
    'Creates a new tab with the same name as the rep
    Worksheets.Add(, Worksheets(Worksheets.Count)).Name = Rep
    
    With Sheets(Rep)
            .Cells(1, 4) = "Name:"
            .Cells(1, 5) = Rep
            .Cells(2, 4) = "Date Paid:"
            .Cells(2, 5) = ReportDate
    End With
    
    'Formats the Name and Date delineators
    With Worksheets(Rep).Range(Sheets(Rep).Cells(1, 4), Sheets(Rep).Cells(2, 4))
        .HorizontalAlignment = xlRight
        .Font.Bold = True
    End With

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Formats the new spreadsheet
    With Sheets(Rep)
        .Range(.Cells(reportrow - 1, 2), .Cells(reportrow - 1, overTypeCol)).Merge
        .Cells(reportrow - 1, 2) = previous_month & " OVERRIDES THAT ARE PAYABLE OR WILL NEED A KILOWATT REPLACEMENT THIS PERIOD "
        .Cells(reportrow, repCol) = "Rep"
        .Cells(reportrow, customerCol) = "Customer"
        .Cells(reportrow, kWCol) = "kW"
        .Cells(reportrow, rateCol) = "Rate/kW"
        .Cells(reportrow, totalCol) = "Total"
        .Cells(reportrow, reasonCol) = "Type"
        .Cells(reportrow, jobIDCol) = "Job ID"
        .Cells(reportrow, statusCol) = "Status"
        .Cells(reportrow, subStatusCol) = "SubStatus"
        .Cells(reportrow, overTypeCol) = "OverrideType"
    End With

    'Formats the column headers
    With Worksheets(Rep).Range(Cells(reportrow - 1, repCol).Address, Cells(reportrow, overTypeCol).Address)
            .HorizontalAlignment = xlCenter
            .Font.Bold = True
            .Interior.Color = RGB(0, 102, 204)
            .Font.Color = RGB(255, 255, 255)
    End With

    For masterrow = 2 To Sheets("Master").Cells(1, 1).End(xlDown).Row + 1
        
        If Sheets("Master").Cells(masterrow, 2) = Rep Then
                'Gives values to the variables to be put into the created worksheet
                With Sheets("Master")
                  subRep = .Cells(masterrow, 4)
                  customer = .Cells(masterrow, 6)
                  kW = .Cells(masterrow, 15)
                  Rate = .Cells(masterrow, 14)
                  Total = .Cells(masterrow, 16)
                  Reason = .Cells(masterrow, 10)
                  JobID = .Cells(masterrow, 7)
                  iStatus = .Cells(masterrow, 11)
                  SubStatus = .Cells(masterrow, 12)
                  overrideType = .Cells(masterrow, 13)
                  grand_total = grand_total + Total
                End With
                
                reportrow = reportrow + 1
                
                'Inputs data into the rep's report
                Sheets(Rep).Cells(reportrow, repCol) = subRep
                Sheets(Rep).Cells(reportrow, customerCol) = customer
                Sheets(Rep).Cells(reportrow, kWCol) = kW
                Sheets(Rep).Cells(reportrow, rateCol) = Rate
                Sheets(Rep).Cells(reportrow, totalCol) = Total
                If Total < 0 Then
                    With Sheets(Rep).Cells(reportrow, totalCol)
                        .Font.Color = RGB(255, 0, 0)
                    End With
                End If
                Sheets(Rep).Cells(reportrow, reasonCol) = Reason
                Sheets(Rep).Cells(reportrow, jobIDCol) = JobID
                Sheets(Rep).Cells(reportrow, statusCol) = iStatus
                Sheets(Rep).Cells(reportrow, subStatusCol) = SubStatus
                Sheets(Rep).Cells(reportrow, overTypeCol) = overrideType
        End If

    Next masterrow
    
    'Sums the totals for each customer into a grand total for the rep
    Worksheets(Rep).Cells(reportrow + 1, 5) = "Total:"
    Worksheets(Rep).Cells(reportrow + 1, 6) = grand_total
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
    
    reportrow = reportrow + 4

    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    'Adjusts column width of the new tab
    With Worksheets(Rep).Range("C1:K1")
        .EntireColumn.AutoFit
    End With

    'Moves the rep row counter to the next rep
    repRow = repRow + 1

    'Creates a workbook for each rep
    'ThisWorkbook.Sheets(Rep).Copy
    'ActiveWorkbook.SaveAs ("C:\users\Rodney\desktop\" & "Payroll Breakdown\" & Rep & ".xlsx")
    'ActiveWorkbook.Close

'Create format for the historical breakdown'
With Sheets(Rep)
        .Range(.Cells(reportrow + 3, 2), .Cells(reportrow + 3, overTypeCol)).Merge
        .Cells(reportrow + 3, 2) = "UPDATED INFORMATION ABOUT EVERY JOB INSIDE OF YOUR DOWNLINE"
        .Cells(reportrow + 4, repCol) = "Rep"
        .Cells(reportrow + 4, customerCol) = "Customer"
        .Cells(reportrow + 4, kWCol) = "kW"
        .Cells(reportrow + 4, rateCol) = "Rate/kW"
        .Cells(reportrow + 4, totalCol) = "Total Paid"
        .Cells(reportrow + 4, reasonCol) = "Date Created"
        .Cells(reportrow + 4, jobIDCol) = "Job ID"
        .Cells(reportrow + 4, statusCol) = "Status"
        .Cells(reportrow + 4, subStatusCol) = "SubStatus"
        .Cells(reportrow + 4, overTypeCol) = "OverrideType"
End With

With Worksheets(Rep).Range(Cells(reportrow + 3, repCol).Address, Cells(reportrow + 4, overTypeCol).Address)
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Interior.Color = RGB(0, 102, 204)
        .Font.Color = RGB(255, 255, 255)
End With
Loop Until Sheets("Reps").Cells(repRow, 1) = ""

''''''''''''''''''TURN OFF SCREEN UPDATING''''''''''''''''''''''
Application.ScreenUpdating = True
End Sub




