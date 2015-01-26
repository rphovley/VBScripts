Function isWorkBookOpen(ByVal FileName As String) As Boolean

     For Each wbk1 In Workbooks
        If (wbk1.Path & "\" & wbk1.Name = FileName) Then
            isWorkBookOpen = True
            Exit For
        End If
    Next
        
End Function

