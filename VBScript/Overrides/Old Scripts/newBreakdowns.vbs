
'MAKE SURE PAYMENTS SHEET IS ORDERED BY ENTRY DATE, THEN OVERRIDEREP NAME, THEN OVERRIDE TYPE, THEN REP NAME BEFORE RUNNING'
Sub overrideBreakdowns()
'MAKE SURE PAYMENTS SHEET IS ORDERED BY ENTRY DATE, THEN OVERRIDEREP NAME, THEN OVERRIDE TYPE, THEN REP NAME BEFORE RUNNING'
Dim i, y, j, firstRow As Integer
    Dim JobID, findID, typeOver, adjustment As String
    Dim CommisionFile, CommisionName, OverRides, OverRideName, overrideMonth As String
    Dim wbk, wbkReps As Workbook
    Dim repsThisMonth As Worksheet
    Dim repRef, newSheet, overRideReps, Master As Worksheet
    Dim repName As String
    Dim hasBeenUsed As Boolean
    CommisionFile = Application.GetOpenFilename()
    OverRides = ThisWorkbook.Name

    CommisionName = convertToName(CommisionFile)
    
    If isWorkBookOpen(CommisionFile) Then
        Set wbkReps = Workbooks(CommisionName)
    Else
        Set wbkReps = Workbooks.Open(CommisionFile)
    End If

    OverRides = ThisWorkbook.Name

    overrideMonth = "September"
    'get access to override sheet'
    Set wbk = Workbooks(OverRides)
    Set repRef = wbk.Sheets("Payments")
    Set overRideReps = wbkReps.Sheets("RepSheetRef")
    Set Master = wbkReps.Sheets("Override Map Master - Master")

    'index for new sheet'
    k = 4
    'index for each override rep'
    j = 1
    'index for this months payments'
    'index for looping through an override rep's reps'
    l = 5

    With repRef

            'find the first month for the current month of overrides'
            firstRow = WorksheetFunction.Match(overrideMonth, .Range("I:I"), 0)
            i = firstRow
            i = 3283
                'Loop through payments to print out breakdowns'
                Do Until IsEmpty(Master.Cells(1, j))
                        Set repsThisMonth = Worksheets.Add()
                        
                        'Go through all of '
                        Do Until IsEmpty(.Range("OverRideRep").Cells(i, 1))
                            Dim isFirst As Boolean
                            Dim repColumn, TypeColumn, RateColumn, CustomerColumn, ReasonColumn, kwColumn, AmountColumn As Integer
                            Dim fixRow As Integer

                            fixRow = 4
                            repColumn = 5
                            TypeColumn = 6
                            RateColumn = 7
                            CustomerColumn = 8
                            ReasonColumn = 9
                            kwColumn = 10
                            AmountColumn = 11
                            isFirst = True
                            hasBeenUsed = False
                            
                            'FORMATTING FOR NEW SHEET'
                            Columns("E:G").Font.Bold = True
                            Columns("H:K").Font.Italic = True
                            Columns("K").NumberFormat = "$#,##0.00;[Red]$#,##0.00"
                            Columns("G").NumberFormat = "$#,##0.00;[Red]$#,##0.00"
                            
                            'Name field'
                            repsThisMonth.Cells(2, 2).Value = "Name"
                            repsThisMonth.Cells(2, 2).Font.Bold = True
                            repsThisMonth.Cells(2, 3).Value = Master.Cells(1, j)

                            'final amount'
                            repsThisMonth.Cells(3, 2).Value = "Final Amount"
                            repsThisMonth.Cells(3, 2).Font.Color = vbWhite
                            repsThisMonth.Range(Cells(3, 2), Cells(3, 3)).Merge
                            repsThisMonth.Cells(3, 2).Interior.Color = Val("&00A6FF&")

                            'Final Amount value formatting'
                            repsThisMonth.Cells(4, 2).Font.Color = vbWhite
                            repsThisMonth.Cells(4, 2).Font.Bold = True
                            repsThisMonth.Range(Cells(4, 2), Cells(4, 3)).Merge
                            repsThisMonth.Cells(4, 2).Interior.Color = Val("&ABABAB&")

                            'Override main section formatting and values'
                            repsThisMonth.Cells(2, repColumn).Value = "Overrides"
                            repsThisMonth.Cells(2, repColumn).Font.Bold = True
                            repsThisMonth.Cells(3, repColumn).Value = "Rep"
                            repsThisMonth.Cells(3, TypeColumn).Value = "Type"
                            repsThisMonth.Cells(3, RateColumn).Value = "Rate"
                            repsThisMonth.Cells(3, CustomerColumn).Value = "Customer"
                            repsThisMonth.Cells(3, ReasonColumn).Value = "Reason"
                            repsThisMonth.Cells(3, kwColumn).Value = "kW"
                            repsThisMonth.Cells(3, AmountColumn).Value = "Amount"
                            repsThisMonth.Range(Cells(2, repColumn), Cells(2, AmountColumn)).Merge
                            repsThisMonth.Range(Cells(2, repColumn), Cells(2, AmountColumn)).Interior.Color = Val("&00A6FF&")
                            repsThisMonth.Range(Cells(2, repColumn), Cells(2, AmountColumn)).Font.Color = vbWhite



                            Dim test1, test2, test3, test4, currentOverrideRepID, currentRepID As String
                            test1 = .Range("OverrideRepID").Cells(i, 1).Value
                            currentOverrideRepID = Master.Cells(3, j).Value
                            test2 = Master.Cells(3, j).Value
                            If .Range("OverrideRepID").Cells(i, 1).Value = currentOverrideRepID Then
                                'print out all the information for this particular override rep'
                                currentRepID = .Range("RepID").Cells(i, 1).Value
                                Do While .Range("RepID").Cells(i, 1).Value = currentRepID
                                        hasBeenUsed = True
                                        If isFirst Then
                                            'Override Rep Name'
                                            repsThisMonth.Cells(k, repColumn).Value = .Range("Rep").Cells(i, 1).Value
                                            'Override Type'
                                            repsThisMonth.Cells(k, TypeColumn).Value = .Range("OverrideType").Cells(i, 1).Value
                                            'Override Rate'
                                            repsThisMonth.Cells(k, RateColumn).Value = .Range("OverrideRate").Cells(i, 1).Value
                                        End If

                                        'override Customer'
                                        repsThisMonth.Cells(k, CustomerColumn).Value = .Range("Customer").Cells(i, 1).Value
                                        'override Reason'
                                        repsThisMonth.Cells(k, ReasonColumn).Value = .Range("Reason").Cells(i, 1).Value
                                        'Override kW'
                                        repsThisMonth.Cells(k, kwColumn).Value = .Range("kW").Cells(i, 1).Value
                                        'override Amount'
                                        repsThisMonth.Cells(k, AmountColumn).Value = .Range("Amount").Cells(i, 1).Value

                                    k = k + 1
                                    i = i + 1
                                    test3 = .Range("OverrideRepID").Cells(i, 1).Value
                                    test4 = .Range("OverrideRepID").Cells(i - 1, 1).Value
                                    'See if this row is the same as the last row, if it is, set isFirst to False'
                                    If .Range("OverrideRepID").Cells(i, 1).Value = .Range("OverrideRepID").Cells(i - 1, 1).Value Then
                                        isFirst = False
                                    Else
                                        isFirst = True
                                        'individual line formatting'
                                        repsThisMonth.Range(Cells(fixRow, repColumn), Cells(k, AmountColumn)).BorderAround xlContinuous
                                        repsThisMonth.Range(Cells(fixRow, repColumn), Cells(k, AmountColumn)).Interior.Color = vbWhite
                                        fixRow = k
                                        
                                    End If
                                Loop
                            
                            'Resize all columns to fit contents'
                            repsThisMonth.Range("E:K").Columns.AutoFit
                            'Final Amount value'
                            repsThisMonth.Cells(4, 2).Value = "Final Amount"
                            
                            Else
                                i = i + 1
                            'INSERT METHOD TO AUTO SEND OUT THIS WORKSHEET'
                            If hasBeenUsed Then
                                Set repsThisMonth = Worksheets.Add()
                            End If
                            hasBeenUsed = False
                            End If

                        Loop

                    'Plus 12 because it needs to skip twelve rows across to the next rep'
                    j = j + 12
                Loop

    End With

End Sub


