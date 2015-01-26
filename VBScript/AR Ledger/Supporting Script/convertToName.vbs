Function convertToName(ByVal Path As String) As String

     For Each wbk1 In Workbooks
        If (wbk1.Path & "\" & wbk1.Name = Path) Then
            convertToName = wbk1.Name
            Exit For
        End If
    Next


End Function

