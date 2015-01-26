'SORT BY JOBID AND OVERRIDEREP BEFORE RUNNING
Sub assignOverID()
'SORT BY JOBID AND OVERRIDEREP BEFORE RUNNING
'BE CAREFUL BEFORE RUNNING THIS check i and overID

    Dim i, j, overID As Integer
    Dim job1, job2, type1, type2 As String
    Dim found As Boolean

    'CHECK THESE
    i = 3642
    overID = 3181
    'CHECK THESE
    
    found = False
    
    With Sheets("Payments")
    
        Do Until IsEmpty(.Cells(i, 2))
            
            j = i - 1

            'What items are similar in the list?
            Do While j > 2
                 
                 job1 = .Cells(i, 5).Value
                 type1 = .Cells(i, 9).Value
                 job2 = .Cells(j, 5).Value
                 type2 = .Cells(j, 9).Value
                'If jobID and overridetype are equal
                
                
                If .Cells(i, 7).Value = .Cells(j, 7).Value And _
                    .Cells(i, 13).Value = .Cells(j, 13).Value And _
                    .Cells(i, 3).Value = .Cells(j, 3).Value Then
                    found = True
                    .Cells(i, 1).Value = .Cells(j, 1).Value
                    Exit Do
                End If
                
                j = j - 1
            Loop
            
            If overID > 500 Then
                    .Cells(8000, 1) = ""
            End If
                
            If found = False Then
                .Cells(i, 1).Value = overID
                    overID = overID + 1
            Else
                found = False
            End If
            i = i + 1
            
            
        Loop
    End With

End Sub