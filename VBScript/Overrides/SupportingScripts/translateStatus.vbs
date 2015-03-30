'is status at first payment?'
Function isReady(ByRef Status As String) As Boolean

    Dim isArray As Variant
    isArray = Array("Submission Hold", "Ready", _
        "Submitted", "Rejected", "Rebate Program Closed", _
        "Design Complete", "Application Complete", _
        "Received")
    
    For Each permitStatus In isArray
    
        If permitStatus = Status Then
            isReady = True
            Exit For
        End If
    Next permitStatus
    
End Function

Function isReadyException(ByRef Status As String) As Boolean

Dim isArray As Variant
    isArray = Array("Submission Hold", "Ready", _
        "Submitted", "Rejected", "Rebate Program Closed", _
        "Design Complete", "Application Complete", _
        "Received", "Site Survey Complete", "Site Survey Scheduled")
    
    For Each permitStatus In isArray
    
        If permitStatus = Status Then
            isReady = True
            Exit For
        End If
    Next permitStatus
    
End Function

'is status a cancelled status'
Function isJobCancelled(ByRef Status As String) As Boolean

    Dim isArray As Variant
    isArray = Array("Customer Uncertain", "Customer Unresponsive", _
        "Job Disqualified", "On Hold")
    
    For Each permitStatus In isArray
    
        If permitStatus = Status Then
            isJobCancelled = True
            Exit For
        End If
    Next permitStatus
    
    
    

'is Status at Backend for new pay structure'
Function isBackend_New(ByRef Status As String, ByRef SubStatus As String) As Boolean

    Dim isArray As Variant
    isArray = Array("Inspection", "Utility", _
        "In Operation", "Closed")

        'Loops through backend statuses that trigger backend'
        For Each arrayStatus In isArray

            'if it is a correct backend status, return true'
            If arrayStatus = Status Then
                isBackend_New = True
                Exit For
            End If

        Next arrayStatus
    
        'This code is only hit if the previous loop didn't return a value
        'The only other situation for a backend is if the substatus = "complete"'
        If SubStatus = "Complete" Then
            isBackend_New = True
        End If

End Function