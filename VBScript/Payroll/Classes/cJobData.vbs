'Class
'Attributes
Private pCustomer, pJobID, pStatus, _
    pSubStatus, pRepEmail As String

Private pDaysSinceCreated As Integer

Private pkW As Double
Private pAmount, pWhatWasPaid, pFirstPaymentAmount, pSecondPaymentAmount, pFinalPaymentAmount As Currency
Private pCreatedDate, pFirstPaymentDate, pSecondPaymentDate, pFinalPaymentDate As Date
Private pIsInstall, pIsDocSigned, pIsFinalContract, pIsCancelled, pIsPaidInFull, pIsBlackListed As Boolean

'Rep Info'
Public Property Get RepID() As Integer
    RepID = pRepID
End Property

Public Property Let RepID(value As Integer)
    pRepID = value
End Property

Public Property Get RepEmail() As String
    RepEmail = pRepEmail
End Property

Public Property Let RepEmail(value As String)
    pRepEmail = value
End Property

Public Property Get RepName() As String
    RepName = pRepName
End Property

Public Property Let RepName(value As String)
    pRepName = value
End Property

Public Property Get PayScaleID() As Integer
    PayScaleID= pPayScaleID
End Property

'Payment Get/Set methods'
Public Property Let IsPaidInFull(value As Boolean)
    pIsPaidInFull = value
End Property

Public Property Get IsPaidInFull() As Boolean
    IsPaidInFull = pIsPaidInFull
End Property

Public Property Get FirstPaymentDate() As Date
    FirstPaymentDate = pFirstPaymentDate
End Property

Public Property Let FirstPaymentDate(value As Date)
    pFirstPaymentDate = value
End Property

Public Property Get SecondPaymentDate() As Date
    SecondPaymentDate = pSecondPaymentDate
End Property

Public Property Let SecondPaymentDate(value As Date)
    pSecondPaymentDate = value
End Property

Public Property Get FinalPaymentDate() As Date
    FinalPaymentDate = pFinalPaymentDate
End Property

Public Property Let FinalPaymentDate(value As Date)
    pFinalPaymentDate = value
End Property

Public Property Get WhatWasPaid() As Currency
    WhatWasPaid = pWhatWasPaid
End Property

Public Property Let WhatWasPaid(value As Currency)
    pWhatWasPaid = value
End Property

Public Sub setWhatWasPaid()
    pWhatWasPaid = pFirstPaymentAmount + pSecondPaymentAmount
End Sub

Public Property Get FirstPaymentAmount() As Currency
    FirstPaymentAmount = pFirstPaymentAmount
End Property

Public Property Let FirstPaymentAmount(value As Currency)
    pFirstPaymentAmount = value
End Property

Public Property Get SecondPaymentAmount() As Currency
    SecondPaymentAmount = pSecondPaymentAmount
End Property

Public Property Let SecondPaymentAmount(value As Currency)
    pSecondPaymentAmount = value
End Property

Public Property Get FinalPaymentAmount() As Currency
    FinalPaymentAmount = pFinalPaymentAmount
End Property

Public Property Let FinalPaymentAmount(value As Currency)
    pFinalPaymentAmount = value
End Property


'Get/Set Methods IsInstall booleans
Public Property Let IsInstall(value As Boolean)
    pIsInstall = value
End Property

Public Property Get IsInstall() As Boolean
    IsInstall = pIsInstall
End Property

Public Sub setIsInstall()
    Dim isArray As Variant
    isArray = Array("Inspection", "Utility", _
        "In Operation", "Closed")

        'Loops through backend statuses that trigger backend'
        For Each arrayStatus In isArray
            pIsInstall = False
            
            'if it is a correct backend status, return true'
            If arrayStatus = Me.Status Then
                pIsInstall = True
                Exit For
            End If
        Next arrayStatus
    
        'This code is only hit if the previous loop didn't return a value
        'The only other situation for a backend is if the substatus = "complete"'
        If Me.SubStatus = "Complete" Then
            pIsInstall = True
        End If

End Sub

'Get/Set Methods IsDocSigned booleans
Public Property Let IsDocSigned(value As Boolean)
    pIsDocSigned = value
End Property

Public Sub setIsDocSigned(ByVal value As String)
    If UCase(value) = "Y" Then
        pIsDocSigned = True
    Else
        pIsDocSigned = False
    End IF
End Sub

Public Property Get IsDocSigned() As Boolean
    IsDocSigned = pIsDocSigned
End Property

'Get/Set Methods IsFinalContract booleans
Public Property Let IsFinalContract(value As Boolean)
    pIsFinalContract = value
End Property

Public Sub setIsFinalContract(ByVal value As String)
    If UCase(value) = "Y" Then
        pIsFinalContract = True
    Else
        pIsFinalContract = False
    End IF
End Sub

Public Property Get IsFinalContract() As Boolean
    IsFinalContract = pIsFinalContract
End Sub

'Get/Set Methods IsCancelled booleans
Public Property Let IsCancelled(value As Boolean)
    pIsCancelled = value
End Property

Public Property Get IsCancelled() As Boolean
    IsCancelled = pIsCancelled
End Property


Public Sub setIsCancelled()
    Dim isArray As Variant
    isArray = Array("Customer Uncertain", "Customer Unresponsive", _
        "Job Disqualified", "On Hold")

    If Me.Status = "Cancelled" Then
        pIsCancelled = True
    Else

        For Each permitStatus In isArray

            If permitStatus = Me.SubStatus Then
                
                pIsCancelled = True
                Exit For

            End If
        Next permitStatus
    End If

End Sub

'Get/Let Methods
Public Property Get CreatedDate() As Date
    CreatedDate = pCreatedDate
End Property

Public Property Let CreatedDate(value As Date)
    pCreatedDate = value
End Property

Public Property Get DaysSinceCreated() As Integer
    DaysSinceCreated = pDaysSinceCreated
End Property

Public Property Let DaysSinceCreated(value As Integer)
    pDaysSinceCreated = value
End Property

Public Sub setDaysSinceCreated()
	pDaysSinceCreated = DateDiff("d",pCreatedDate, Now())
End Sub

Public Property Get Amount() As Currency
    Amount = pAmount
End Property

Public Property Let Amount(value As Currency)
    pAmount = value
End Property

Public Property Get kW() As Double
    kW = pkW
End Property

Public Property Let kW(value As Double)
    pkW = value
End Property

Public Property Get Status() As String
    Status = pStatus
End Property

Public Property Let Status(value As String)
    pStatus = value
End Property

Public Property Get SubStatus() As String
    SubStatus = pSubStatus
End Property

Public Property Let SubStatus(value As String)
    pSubStatus = value
End Property

Public Property Get JobID() As String
    JobID = pJobID
End Property

Public Property Let JobID(value As String)
    pJobID = value
End Property

Public Property Get Customer() As String
    Customer = pCustomer
End Property

Public Property Let Customer(value As String)
    pCustomer = value
End Property



''''''''''''''''''''''''METHODS''''''''''''''''''''''
