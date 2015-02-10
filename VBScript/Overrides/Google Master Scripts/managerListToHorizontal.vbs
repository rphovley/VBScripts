'transform manager list into horizontal merged list'
Private Sub transform()

Dim i, j As Integer
Dim from As Worksheet
Dim Master As Workbook

    Set Master = Workbooks(ThisWorkbook.Name)

    Set from = Workbooks(ThisWorkbook.Name).Sheets("RepsEmail")
    
    i = 1
    j = 1
    With Sheets("Sheet1")
        
        Do Until IsEmpty(from.Cells(i, 9))
        
            .Range(Cells(1, j), Cells(1, j + 11)).Merge
            .Range(Cells(2, j), Cells(2, j + 11)).Merge
            
            .Cells(1, j).Value = from.Cells(i, 9).Value
            .Cells(2, j).Value = from.Cells(i, 10).Value
            
            j = j + 12
            i = i + 1
        Loop
    End With
End Sub
