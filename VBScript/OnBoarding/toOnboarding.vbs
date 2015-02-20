Sub AccountInfo()

Dim CurrentRep As cOnboardingList
Dim PrintRep As cOnboardingList
Dim arrayRep() As cOnboardingList

Dim Master As Workbook
Dim Onboarding As Workbook
Dim OnboardingList As Worksheet
Dim DecemberMap As Worksheet

Dim TotalReps As Long
Dim usedCol As Long
Dim rep_row As Long
Dim report_row As Long
Dim y As Long

Dim RCCounter, MCounter, RGCounter, DCounter As Integer

Set Master = Workbooks("January Override Master.xlsm")
Set DecemberMap = Master.Sheets("January 2015 Map")
Set Onboarding = Workbooks("Onboarding.xlsm")
Set OnboardingList = Onboarding.Sheets("OnboardingList")

report_row = 3
usedCol = DecemberMap.UsedRange.Columns.Count
TotalReps = DecemberMap.UsedRange.Rows.Count

'Puts the data in the Class
    For rep_row = 3 To TotalReps
        Set CurrentRep = New cOnboardingList
        ReportOffset = 1
        RCCounter = 0
        MCounter = 0
        RGCounter = 0
        DCounter = 0
        y = 3
        
        With DecemberMap
        
            CurrentRep.RepName = .Range(.Cells(rep_row, 1), .Cells(rep_row, 1)).value
            CurrentRep.RepID = .Range(.Cells(rep_row, 2), .Cells(rep_row, 2)).value
            
            y = 3
            
            Do Until .Cells(rep_row, y) = ""
                If .Cells(rep_row, y) = "RC" Then
                    CurrentRep.Recruiter = .Range(.Cells(rep_row, y + 1), .Cells(rep_row, y + 1)).value
                    CurrentRep.RecruiterID = .Range(.Cells(rep_row, y + 2), .Cells(rep_row, y + 2)).value
                    CurrentRep.RecruiterRate = .Range(.Cells(rep_row, y + 3), .Cells(rep_row, y + 3)).value
                    
                    OnboardingList.Cells(report_row + RCCounter, 4) = CurrentRep.Recruiter()
                    OnboardingList.Cells(report_row + RCCounter, 5) = CurrentRep.RecruiterID()
                    OnboardingList.Cells(report_row + RCCounter, 6) = CurrentRep.RecruiterRate()
                    
                    RCCounter = RCCounter + 1
                ElseIf .Cells(rep_row, y) = "M" Then
                    CurrentRep.Manager = .Range(.Cells(rep_row, y + 1), .Cells(rep_row, y + 1)).value
                    CurrentRep.ManagerID = .Range(.Cells(rep_row, y + 2), .Cells(rep_row, y + 2)).value
                    CurrentRep.ManagerRate = .Range(.Cells(rep_row, y + 3), .Cells(rep_row, y + 3)).value
                    
                    OnboardingList.Cells(report_row + MCounter, 8) = CurrentRep.Manager()
                    OnboardingList.Cells(report_row + MCounter, 9) = CurrentRep.ManagerID()
                    OnboardingList.Cells(report_row + MCounter, 10) = CurrentRep.ManagerRate()
                    
                    MCounter = MCounter + 1
                ElseIf .Cells(rep_row, y) = "RG" Then
                    CurrentRep.Regional = .Range(.Cells(rep_row, y + 1), .Cells(rep_row, y + 1)).value
                    CurrentRep.RegionalID = .Range(.Cells(rep_row, y + 2), .Cells(rep_row, y + 2)).value
                    CurrentRep.RegionalRate = .Range(.Cells(rep_row, y + 3), .Cells(rep_row, y + 3)).value
                    
                    OnboardingList.Cells(report_row + RGCounter, 12) = CurrentRep.Regional()
                    OnboardingList.Cells(report_row + RGCounter, 13) = CurrentRep.RegionalID()
                    OnboardingList.Cells(report_row + RGCounter, 14) = CurrentRep.RegionalRate()
                    
                    RGCounter = RGCounter + 1
                ElseIf .Cells(rep_row, y) = "D" Then
                    CurrentRep.DVP = .Range(.Cells(rep_row, y + 1), .Cells(rep_row, y + 1)).value
                    CurrentRep.DVPID = .Range(.Cells(rep_row, y + 2), .Cells(rep_row, y + 2)).value
                    CurrentRep.DVPRate = .Range(.Cells(rep_row, y + 3), .Cells(rep_row, y + 3)).value
                    
                    OnboardingList.Cells(report_row + DCounter, 16) = CurrentRep.DVP()
                    OnboardingList.Cells(report_row + DCounter, 17) = CurrentRep.DVPID()
                    OnboardingList.Cells(report_row + DCounter, 18) = CurrentRep.DVPRate()
                    
                    DCounter = DCounter + 1
                End If
                y = y + 4
            
            If Application.Max(RCCounter, MCounter, RGCounter, DCounter) > 0 Then
                ReportOffset = Application.Max(RCCounter, MCounter, RGCounter, DCounter)
            End If
        
                 With OnboardingList
                    If .Cells(report_row + RCCounter, 5) = 0 Then
                        .Cells(report_row + RCCounter, 5) = ""
                        .Cells(report_row + RCCounter, 6) = ""
                    End If
                    If .Cells(report_row + MCounter, 9) = 0 Then
                        .Cells(report_row + MCounter, 9) = ""
                        .Cells(report_row + MCounter, 10) = ""
                    End If
                    If .Cells(report_row + RGCounter, 13) = 0 Then
                        .Cells(report_row + RGCounter, 13) = ""
                        .Cells(report_row + RGCounter, 14) = ""
                    End If
                    If .Cells(report_row + DCounter, 17) = 0 Then
                        .Cells(report_row + DCounter, 17) = ""
                        .Cells(report_row + DCounter, 18) = ""
                    End If
                    
                End With
            Loop
        End With
        
        ReDim Preserve arrayRep(3 To rep_row)
        Set arrayRep(rep_row) = CurrentRep
        Set CurrentRep = Nothing
    
    'print out everything to desired location
            Set PrintRep = arrayRep(rep_row)
        
            With OnboardingList
                .Cells(report_row, 1) = PrintRep.RepName()
                .Cells(report_row, 2) = PrintRep.RepID()
                
                .Range(.Cells(report_row, 1), .Cells(report_row, 18)).Borders(xlEdgeTop).LineStyle = xlContinuous
            End With
            
        report_row = report_row + ReportOffset
        
            Set PrintRep = Nothing
            
    Next rep_row

'opens up more memory in Excel
    For rep_row = 3 To TotalReps
        Set arrayRep(rep_row) = Nothing
    Next rep_row
    
    Set DecemberMap = Nothing
    
End Sub

