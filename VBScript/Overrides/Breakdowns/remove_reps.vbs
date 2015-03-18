Sub removeReps()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''INITIALIZE VARIABLES''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''      
    Dim repCol As Integer

        repCol   = 2


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''GET AND SET VALUES''''''''''''''''''''''''''''
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Dim repDataSize As Integer
    Dim doNotPayRep As String

        With WorkSheets("Do Not Pay")
            repDataSize = .Cells(1,1).End(xlDown).Row  
        End With
      For inputRow = 2 To repDataSize
            With WorkSheets("Do Not Pay")
                doNotPayRep = .Cells(inputRow,1).value 
            End With

            With WorkSheets("Master")

                For masterRow = 2 To .Cells(1,1).End(xlDown).Row
                    If .Cells(masterRow, repCol).value = doNotPayRep Then
                        .Rows(masterRow).EntireRow.Delete
                    End If
                Next masterRow
            End With
      Next inputRow
        
End Sub