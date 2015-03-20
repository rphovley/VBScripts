'Class
'Attributes

Private pCustomer, pJobID, pStatus, _
    pSubStatus, pRepEmail, pStates As String
Private pDaysSinceCreated, pFirstPaymentRow, pSecondPaymentRow, pInstallRow As Integer
Private pkW As Double
Private pAmount, pWhatWasPaid, pFirstPaymentAmount, pSecondPaymentAmount, pFinalPaymentAmount, pClawbackAmount As Currency
private pThisWeekFirstPayment, pThisWeekSecondPayment, pThisWeekFinalPayment, pThisWeekCancelled As Currency
Private pCreatedDate, pFirstPaymentDate, pSecondPaymentDate, pFinalPaymentDate, pDateOfClawback As Date
Private pIsInstall, pIsDocSigned, pIsSurveyComplete, pIsFinalContract, pIsCancelled, pIsPaidInFull, pIsBlackListed As Boolean

Public Property Get FirstPaymentRow() As Integer
    FirstPaymentRow = pFirstPaymentRow
End Property

Public Property Let FirstPaymentRow(value As Integer)
    pFirstPaymentRow = value
End Property

Public Property Get SecondPaymentRow() As Integer
    SecondPaymentRow = pSecondPaymentRow
End Property

Public Property Let SecondPaymentRow(value As Integer)
    pSecondPaymentRow = value
End Property

Public Property Get InstallRow() As Integer
    InstallRow = pInstallRow
End Property

Public Property Let InstallRow(value As Integer)
    pInstallRow = value
End Property

Public Property Get DateOfClawback() As Date
    DateOfClawback = pDateOfClawback
End Property

Public Property Let DateOfClawback(value As Date)
    pDateOfClawback = value
End Property

Public Property Get ClawbackAmount() As Currency
    ClawbackAmount = pClawbackAmount
End Property

Public Property Let ClawbackAmount(value As Currency)
    pClawbackAmount = value
End Property

Public Property Get States() As String
    States = pStates
End Property

Public Property Let States(value As String)
    pStates = value
End Property

Public Property Get ThisWeekFirstPayment() As Currency
    ThisWeekFirstPayment = pThisWeekFirstPayment
End Property

Public Property Let ThisWeekFirstPayment(value As Currency)
    pThisWeekFirstPayment = value
End Property

Public Property Get ThisWeekSecondPayment() As Currency
    ThisWeekSecondPayment = pThisWeekSecondPayment
End Property

Public Property Let ThisWeekSecondPayment(value As Currency)
    pThisWeekSecondPayment = value
End Property

Public Property Get ThisWeekFinalPayment() As Currency
    ThisWeekFinalPayment = pThisWeekFinalPayment
End Property

Public Property Let ThisWeekFinalPayment(value As Currency)
    pThisWeekFinalPayment= value
End Property

Public Property Get ThisWeekCancelled() As Currency
    ThisWeekCancelled = pThisWeekCancelled
End Property

Public Property Let ThisWeekCancelled(value As Currency)
    pThisWeekCancelled= value
End Property

'Rep Info'
Public Property Get RepEmail() As String
    RepEmail = pRepEmail
End Property

Public Property Let RepEmail(value As String)
    pRepEmail = value
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
    pWhatWasPaid = pFirstPaymentAmount + pSecondPaymentAmount + pFinalPaymentAmount
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
Public Property Let IsSurveyComplete(value As Boolean)
    pIsSurveyComplete = value
End Property

Public Property Get IsSurveyComplete() As Boolean
    IsSurveyComplete = pIsSurveyComplete
End Property

Public Sub setIsSurveyComplete()
    Dim isArray, statusArray, stateArray As Variant
    isArray = Array("Site Survey Complete", "Design Complete", "Application Complete", _
        "Submitted", "Rejected", "Received")
    statusArray = Array("Sales", "Permit")
    stateArray  = Array("NY", "NJ", "DE", "MD", "CT", "MA")
    Dim isInclementState As Boolean

    For Each sState In stateArray
        If Me.States = sState Then
            isInclementState = True
        End IF
    Next
        'Loops through backend statuses that trigger backend'
            If Me.Status = "Permit" Then
                For Each arrayStatus In isArray
                    'if it is a correct backend status, return true'
                    If NOT isInclementState AND arrayStatus = Me.SubStatus Then
                        Me.IsSurveyComplete = True
                        Exit For
                    ElseIf isInclementState AND arrayStatus = Me.SubStatus OR "Site Survey Scheduled" = Me.Substatus Then
                        Me.IsSurveyComplete = True
                        Exit For
                    End If

                Next arrayStatus

            ElseIf Me.Status = "Sales" Then
                Me.IsSurveyComplete = False
            ElseIf NOT Me.IsCancelled Then
                Me.IsSurveyComplete = True
            End If
End Sub

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
    If UCase(value) <> "N" Then
        pIsDocSigned = True
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
    If value <> "" Then
        pIsFinalContract = True
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
