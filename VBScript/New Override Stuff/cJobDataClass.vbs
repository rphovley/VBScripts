'Class
'Attributes
Private pCustomer, pJobID, pStatus, _
    pSubStatus, pRepEmail, pRepName As String

Private pkW As Double
Private pAmount As Currency
Private pRepID As Integer
Private pCreatedDate As Date
Private pIsInstall, pIsDocSigned, pIsFinalContract, pIsCancelled As Boolean

'Get Methods booleans
Public Property Get IsInstall() As Date
    createdDate = pCreatedDate
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
Public Sub isInstall()
    
    
End Sub