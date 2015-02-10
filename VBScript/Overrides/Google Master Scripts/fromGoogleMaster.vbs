'Check i before running
Sub fromGoogleMaster()
'Check i before running

    Dim i, j, k As Integer
    Dim googleSheet, Master As Workbook
    Dim map, from As Worksheet
    Dim CommisionFile, CommisionName As String
    
    'CHECK THIS
    i = 1
    'CHECK THIS
    
    k = 6
    
    CommisionFile = Application.GetOpenFilename()
    CommisionName = convertToName(CommisionFile)
    
    If isWorkBookOpen(CommisionFile) Then
        Set googleSheet = Workbooks(CommisionName)
    Else
        Set googleSheet = Workbooks.Open(CommisionFile)
    End If
       
    Set Master = Workbooks(ThisWorkbook.Name)
    Set map = Master.Sheets("x")
    Set from = googleSheet.Sheets("Master")
    
    With map
        
        'loop across sheet to each manager
        Do Until IsEmpty(from.Cells(1, i))
            k = 6
            For j = i To i + 11
                k = 6
                Do Until IsEmpty(from.Cells(k, j))
                    Dim repRow, endColumn As Integer
                    Dim colType, colName, colID, colRate As Integer
                    colType = 0
                    colName = 1
                    colID = 2
                    colRate = 3
                    
                    'find where to transfer the data from the google spreadsheet
                    repRow = WorksheetFunction.Match(from.Cells(k, j + 1).Value, .Range("B:B"), 0)
                    endColumn = .Cells(repRow, 1).End(xlToRight).Column + 1
                    'Name (Uses translation across columns to find the override rep's Name
                    .Cells(repRow, endColumn + colName).Value = from.Cells(1, i).Value
                    
                    'overrideRepID
                    .Cells(repRow, endColumn + colID).Value = from.Cells(4, i).Value
                    'Rate
                    .Cells(repRow, endColumn + colRate).Value = from.Cells(k, j + 2).Value
                    If Not IsNumeric(.Cells(repRow, endColumn + colRate).Value) Then
                        MsgBox "The Rate at line " & repRow & " was not a number", vbOKOnly, "ERROR"
                    End If

                    'which level of override is it?
                    If from.Cells(5, j).Value = "Recruit" Then
                        'Type
                        .Cells(repRow, endColumn + colType).Value = "RC"
                        
                    ElseIf from.Cells(5, j).Value = "Managed" Then
                        'Type
                        .Cells(repRow, endColumn + colType).Value = "M"
                        
                    ElseIf from.Cells(5, j).Value = "Regional" Then
                        'Type
                        .Cells(repRow, endColumn + colType).Value = "RG"
                        
                    ElseIf from.Cells(5, j).Value = "DVP" Then
                        'Type
                        .Cells(repRow, endColumn + colType).Value = "D"
                        
                    End If
                    
                k = k + 1
                Loop
                j = j + 2
            Next j
        i = i + 12
        Loop
    
    End With
End Sub
