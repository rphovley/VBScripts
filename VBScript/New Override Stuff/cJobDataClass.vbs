'Class
'Attributes
Private pCustomer, pJobID, pStatus, _
    pSubStatus, pRepEmail, pRepName As String

Private pkW As Double
Private pAmount As Currency
Private pRepID As Integer
Private pCreatedDate As Date
Private pIsInstall, pIsDocSigned, pIsFinalContract, pIsCancelled As Boolean

'Get/Set Methods IsInstall booleans
Public Property Get IsInstall() As Date
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
Public Property Get IsDocSigned() As Date
    IsDocSigned = pIsDocSigned
End Property

Public 
'Get/Set Methods IsFinalContract booleans
Public Property Get IsFinalContract() As Date
    IsInstall = pIsInstall
End Property

'Get/Set Methods IsCancelled booleans
Public Property Get IsCancelled() As Date
    IsInstall = pIsInstall
End Property


'Get/Let Methods
Public Property Get CreatedDate() As Date
    CreatedDate = pCreatedDate
End Property

Public Property Let CreatedDate(value As Date)
    pCreatedDate = value
End Property

Public Property Get RepID() As Integer
    RepID = pRepID
End Property

Public Property Let RepID(value As Integer)
    pRepID = value
End Property

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

Public Property Get RepName() As String
    RepName = pRepName
End Property

Public Property Let RepName(value As String)
    pRepName = value
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

Public Property Get RepEmail() As String
    RepEmail = pRepEmail
End Property

Public Property Let RepEmail(value As String)
    pRepEmail = value
End Property

''''''''''''''''''''''''METHODS''''''''''''''''''''''
